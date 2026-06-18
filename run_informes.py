#!/usr/bin/env python3
"""
run_informes.py — Orquestador Siigo → Supabase (diseño desagregado).

Tablas que alimenta:
  • documentos    — una fila por FV o NC (facturas y notas crédito)
  • recibos_caja  — una fila por línea de pago en un RC (recibo de caja)

Variables de entorno requeridas:
  SIIGO_USERNAME    SIIGO_ACCESS_KEY    SIIGO_PARTNER_ID    SUPABASE_DB_URL

Modos de ejecución:
  python3 run_informes.py                          # mes en curso (día 1 → ayer)
  python3 run_informes.py --mes-pasado             # mes anterior completo
  python3 run_informes.py --from-date 2026-01-01 --to-date 2026-06-17

Opciones:
  --solo documentos   # solo sube FV + NC
  --solo recibos      # solo sube RC
  --timeout N         # segundos por página HTTP (default 90)

Comportamiento con rangos multi-mes:
  • documentos: itera mes a mes (límite de la API de Siigo)
  • recibos:    una sola descarga desde from_date hasta hoy
                (un RC de junio puede pagar una FV de enero)

Códigos de salida:
  0  todo ok   1  error parcial   2  error de configuración

Contingencias:
  • Reintentos 429/5xx/red: transparentes desde siigo_core
  • DF/ND no disponibles en la API → nunca fallan, solo no aparecen
  • Phantom IVA: total_cop usa doc.total directamente (no recalcula)
  • FV en USD: tasa de cambio desde el campo currency de la factura
  • Documentos anulados: excluidos automáticamente
  • Clientes sin detalle: NIT como nombre, no falla
"""
from __future__ import annotations

import argparse
import json
import sys
import traceback
from datetime import date, datetime, timedelta, timezone
from typing import Any

from siigo_core import (
    SiigoApiError,
    build_siigo_config,
    fetch_credit_notes,
    fetch_customer_details,
    fetch_users_map,
    safe_text,
)
from ventas_cliente_report import fetch_invoices
from cartera_report import fetch_cost_centers, fetch_cartera_invoices, build_cartera_rows
from documentos_report import (
    build_documentos_rows,
    build_recibos_rows,
    fetch_vouchers_raw,
)
from supabase_writer import upsert_documentos, upsert_recibos

_KNOWN_LIMITS = [
    "DF (Débito Facturación) y ND (Nota Débito) no accesibles vía API pública — no aparecen.",
    "FV en USD: tasa de cambio fija al momento de emisión de la factura.",
    "Phantom IVA (precio=0, impuesto>0): total_cop toma doc.total de Siigo directamente.",
    "FV sin vendedor (cliente 900276962, moneda USD): Siigo no asigna vendedor a estas facturas.",
]


# ── Períodos ──────────────────────────────────────────────────────────────────

def _period_mes_actual() -> tuple[date, date]:
    today     = date.today()
    yesterday = today - timedelta(days=1)
    first_day = today.replace(day=1)
    if yesterday < first_day:
        last_prev = today - timedelta(days=1)
        return last_prev.replace(day=1), last_prev
    return first_day, yesterday


def _period_mes_pasado() -> tuple[date, date]:
    first_this = date.today().replace(day=1)
    last_prev  = first_this - timedelta(days=1)
    return last_prev.replace(day=1), last_prev


def _months_in_range(from_date: date, to_date: date) -> list[tuple[date, date]]:
    """Divide el rango en sub-períodos mensuales para respetar límites de la API."""
    months = []
    cursor = from_date.replace(day=1)
    while cursor <= to_date:
        first = max(cursor, from_date)
        if cursor.month == 12:
            last = cursor.replace(year=cursor.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            last = cursor.replace(month=cursor.month + 1, day=1) - timedelta(days=1)
        last = min(last, to_date)
        months.append((first, last))
        if cursor.month == 12:
            cursor = cursor.replace(year=cursor.year + 1, month=1, day=1)
        else:
            cursor = cursor.replace(month=cursor.month + 1, day=1)
    return months


def _period_tag(from_date: date, to_date: date) -> str:
    if from_date.year == to_date.year and from_date.month == to_date.month:
        return from_date.strftime("%Y%m")
    return f"{from_date.strftime('%Y%m')}_{to_date.strftime('%Y%m')}"


# ── Lookups compartidos ───────────────────────────────────────────────────────

def _fetch_lookups(config: Any, docs: list[dict], timeout: int):
    """Descarga cost_centers, users y customer_details para una lista de documentos."""
    cost_centers_map = fetch_cost_centers(config, timeout=timeout)
    users_map, _     = fetch_users_map(config, timeout=timeout)

    customer_hints: dict[str, dict[str, str]] = {}
    for doc in docs:
        cust = doc.get("customer")
        if isinstance(cust, dict):
            cid = safe_text(cust.get("id"))
            if cid:
                customer_hints[cid] = {"identification": safe_text(cust.get("identification"))}

    customers_map, failures = fetch_customer_details(config, customer_hints, timeout=timeout)
    if failures:
        print(f"  ⚠  {len(failures)} cliente(s) sin detalle — se usará NIT como nombre.", file=sys.stderr)

    return cost_centers_map, users_map, customers_map


# ── Operaciones individuales ──────────────────────────────────────────────────

def _run_documentos(
    config: Any,
    from_date: date,
    to_date: date,
    timeout: int,
) -> dict[str, Any]:
    """
    Descarga FV + NC del período y hace upsert en la tabla documentos.
    Itera mes a mes si el rango abarca varios meses.
    """
    months   = _months_in_range(from_date, to_date)
    total_fv = total_nc = total_rows = 0

    for m_from, m_to in months:
        print(f"\n[Documentos]  {m_from} → {m_to}")

        print("  Descargando facturas (FV)…")
        invoices = fetch_invoices(config, m_from, m_to, timeout=timeout)
        print(f"  {len(invoices)} FV.")

        print("  Descargando notas crédito (NC)…")
        credit_notes = fetch_credit_notes(config, m_from, m_to, timeout=timeout)
        print(f"  {len(credit_notes)} NC.")

        cost_centers_map, users_map, customers_map = _fetch_lookups(
            config, invoices + credit_notes, timeout
        )

        rows = build_documentos_rows(
            invoices, credit_notes, customers_map, users_map, cost_centers_map
        )
        n = upsert_documentos(rows)
        print(f"  ✓ {n} filas en documentos.")

        total_fv   += len(invoices)
        total_nc   += len(credit_notes)
        total_rows += n

    return {
        "informe":  "documentos",
        "estado":   "ok",
        "filas":    total_rows,
        "fv":       total_fv,
        "nc":       total_nc,
        "meses":    len(months),
    }


def _run_recibos(
    config: Any,
    from_date: date,
    timeout: int,
) -> dict[str, Any]:
    """
    Descarga todos los RC desde from_date hasta hoy y hace upsert en recibos_caja.
    No itera mes a mes: un RC de junio puede pagar una FV de enero.
    """
    print(f"\n[Recibos de caja]  {from_date} → {date.today()}")

    print("  Descargando vouchers (RC)…")
    vouchers = fetch_vouchers_raw(config, from_date, date.today(), timeout=timeout)
    print(f"  {len(vouchers)} vouchers.")

    users_map, _ = fetch_users_map(config, timeout=timeout)

    customer_hints: dict[str, dict[str, str]] = {}
    for v in vouchers:
        cust = v.get("customer") or {}
        cid  = safe_text(cust.get("id"))
        if cid:
            customer_hints[cid] = {"identification": safe_text(cust.get("identification"))}

    customers_map, failures = fetch_customer_details(config, customer_hints, timeout=timeout)
    if failures:
        print(f"  ⚠  {len(failures)} cliente(s) sin detalle.", file=sys.stderr)

    rows = build_recibos_rows(vouchers, users_map, customers_map)
    n    = upsert_recibos(rows)
    print(f"  ✓ {n} filas en recibos_caja.")

    return {
        "informe":   "recibos_caja",
        "estado":    "ok",
        "filas":     n,
        "vouchers":  len(vouchers),
    }


def _run_cartera_update(config: Any, timeout: int) -> dict[str, Any]:
    """
    Actualiza el balance_cop de todas las FV que aún tienen saldo > 0.
    También inserta facturas antiguas que no estén en la tabla todavía.
    """
    as_of     = date.today()
    from_date = as_of - timedelta(days=365 * 3)

    print(f"\n[Actualización balances cartera]  corte al {as_of}")

    cost_centers_map = fetch_cost_centers(config, timeout=timeout)
    users_map, _     = fetch_users_map(config, timeout=timeout)

    print("  Descargando facturas con saldo > 0…")
    invoices = fetch_cartera_invoices(config, from_date, as_of, timeout=timeout)
    print(f"  {len(invoices)} facturas con saldo pendiente.")

    customer_hints: dict[str, dict[str, str]] = {}
    for inv in invoices:
        cust = inv.get("customer")
        if isinstance(cust, dict):
            cid = safe_text(cust.get("id"))
            if cid:
                customer_hints[cid] = {"identification": safe_text(cust.get("identification"))}

    customers_map, failures = fetch_customer_details(config, customer_hints, timeout=timeout)
    if failures:
        print(f"  ⚠  {len(failures)} clientes sin detalle.", file=sys.stderr)

    # Reutilizamos build_documentos_rows: solo FV, sin NC
    rows = build_documentos_rows(
        invoices, [], customers_map, users_map, cost_centers_map
    )
    n = upsert_documentos(rows)
    print(f"  ✓ {n} filas actualizadas/insertadas en documentos.")

    total      = sum(r["balance_cop"] for r in rows)
    vencido    = sum(
        r["Vencido 1 a 30"] + r["Vencido 31 a 60"] +
        r["Vencido 61 a 90"] + r["Vencido más de 91"]
        for r in build_cartera_rows(invoices, cost_centers_map, users_map, customers_map, as_of)
    )

    return {
        "informe":        "cartera_update",
        "estado":         "ok",
        "filas":          n,
        "corte":          as_of.isoformat(),
        "total_cartera":  round(total, 2),
        "saldo_vencido":  round(vencido, 2),
    }


# ── CLI ───────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Orquestador Siigo → Supabase (tablas documentos + recibos_caja).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python3 run_informes.py                              # mes actual\n"
            "  python3 run_informes.py --mes-pasado                 # mes anterior\n"
            "  python3 run_informes.py --from-date 2026-01-01 --to-date 2026-06-17\n"
            "  python3 run_informes.py --solo documentos\n"
        ),
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--mes-actual",  action="store_true", help="Mes en curso (default)")
    mode.add_argument("--mes-pasado",  action="store_true", help="Mes anterior completo")
    p.add_argument("--from-date", metavar="YYYY-MM-DD")
    p.add_argument("--to-date",   metavar="YYYY-MM-DD")
    p.add_argument(
        "--solo",
        choices=["documentos", "recibos", "cartera"],
        help="Ejecuta solo una operación",
    )
    p.add_argument("--timeout", type=int, default=90,
                   help="Timeout HTTP por página (default: 90s)")
    return p.parse_args()


def main() -> int:
    args = _parse_args()

    # ── Período ───────────────────────────────────────────────────────────────
    if args.from_date or args.to_date:
        if not (args.from_date and args.to_date):
            print("ERROR: --from-date y --to-date deben usarse juntos.", file=sys.stderr)
            return 2
        try:
            from_date = date.fromisoformat(args.from_date)
            to_date   = date.fromisoformat(args.to_date)
        except ValueError as exc:
            print(f"ERROR: fecha inválida — {exc}", file=sys.stderr)
            return 2
        if from_date > to_date:
            print("ERROR: --from-date no puede ser posterior a --to-date.", file=sys.stderr)
            return 2
    elif args.mes_pasado:
        from_date, to_date = _period_mes_pasado()
    else:
        from_date, to_date = _period_mes_actual()

    tag = _period_tag(from_date, to_date)

    print(f"\n{'='*62}")
    print(f"  SIIGO → SUPABASE  |  {from_date} → {to_date}  |  {tag}")
    print(f"{'='*62}")

    # ── Config ────────────────────────────────────────────────────────────────
    try:
        config = build_siigo_config()
    except (SiigoApiError, RuntimeError) as exc:
        print(f"\nERROR de configuración: {exc}", file=sys.stderr)
        print(
            "Verifica SIIGO_USERNAME, SIIGO_ACCESS_KEY, SIIGO_PARTNER_ID y SUPABASE_DB_URL.",
            file=sys.stderr,
        )
        return 2

    results: list[dict[str, Any]] = []

    # ── Documentos (FV + NC) ──────────────────────────────────────────────────
    if args.solo in (None, "documentos"):
        try:
            results.append(_run_documentos(config, from_date, to_date, args.timeout))
        except Exception as exc:
            print(f"\n  ✗ Documentos FALLÓ: {exc}", file=sys.stderr)
            results.append({
                "informe": "documentos", "estado": "error",
                "error": str(exc), "traceback": traceback.format_exc(),
            })

    # ── Recibos de caja (RC) ──────────────────────────────────────────────────
    if args.solo in (None, "recibos"):
        try:
            results.append(_run_recibos(config, from_date, args.timeout))
        except Exception as exc:
            print(f"\n  ✗ Recibos FALLÓ: {exc}", file=sys.stderr)
            results.append({
                "informe": "recibos_caja", "estado": "error",
                "error": str(exc), "traceback": traceback.format_exc(),
            })

    # ── Actualización balances cartera ────────────────────────────────────────
    if args.solo in (None, "cartera"):
        try:
            results.append(_run_cartera_update(config, args.timeout))
        except Exception as exc:
            print(f"\n  ✗ Cartera update FALLÓ: {exc}", file=sys.stderr)
            results.append({
                "informe": "cartera_update", "estado": "error",
                "error": str(exc), "traceback": traceback.format_exc(),
            })

    # ── Resumen ───────────────────────────────────────────────────────────────
    ok_count  = sum(1 for r in results if r["estado"] == "ok")
    err_count = sum(1 for r in results if r["estado"] == "error")

    print(f"\n{'─'*62}")
    for r in results:
        if r["estado"] == "ok":
            print(f"  ✓  {r['informe']:<28} {r.get('filas', '?')} filas")
        else:
            print(f"  ✗  {r['informe']:<28} ERROR: {r['error']}", file=sys.stderr)

    print(f"\n  {ok_count}/{len(results)} operaciones completadas")
    if err_count:
        print(f"  ⚠  {err_count} error(es). Re-ejecuta con --solo <nombre>.", file=sys.stderr)
    print(f"{'─'*62}\n")

    print(json.dumps({
        "generado_en": datetime.now(timezone.utc).isoformat(),
        "periodo":     {"desde": from_date.isoformat(), "hasta": to_date.isoformat()},
        "operaciones": results,
        "limitaciones_conocidas": _KNOWN_LIMITS,
    }, ensure_ascii=False, indent=2))

    return 0 if err_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
