#!/usr/bin/env python3
"""
run_informes.py — Orquestador de informes Siigo → Supabase.

Descarga datos de la API Siigo y los sube directamente a las tablas de Supabase:
  • ventas_por_cliente        — agrupado por cliente para el período
  • cartera                   — snapshot al día de hoy
  • comisiones_detalle        — facturas cobradas vía recibo de caja
  • comisiones_pendientes     — facturas sin cobrar (snapshot al día de hoy)

Variables de entorno requeridas:
  SIIGO_USERNAME    — usuario de la cuenta Siigo
  SIIGO_ACCESS_KEY  — llave de acceso Siigo
  SIIGO_PARTNER_ID  — Partner-Id del header
  SUPABASE_DB_URL   — cadena de conexión PostgreSQL de Supabase

Modos de fecha (aplican a ventas y comisiones; cartera siempre corre al día de hoy):
  python3 run_informes.py                          # mes en curso (día 1 → ayer) [default]
  python3 run_informes.py --mes-pasado             # mes anterior completo
  python3 run_informes.py --from-date 2026-01-01 --to-date 2026-06-30

Opciones adicionales:
  --solo ventas       # solo sube ventas por cliente
  --solo comisiones   # solo sube comisiones
  --solo cartera      # solo sube cartera
  --timeout N         # segundos por página HTTP (default 90)

Códigos de salida:
  0  todos los informes subidos correctamente
  1  al menos un informe falló
  2  error de configuración (credenciales faltantes, etc.)

Contingencias incorporadas:
  • Reintentos automáticos ante 429/5xx/errores de red (en siigo_core)
  • Documentos DF no disponibles vía API → advertencia en log, no falla
  • Phantom IVA (precio=0, impuesto>0) → se usa doc.total de Siigo directamente
  • Facturas en USD → conversión por tasa de cambio almacenada en la factura
  • Documentos anulados → se excluyen automáticamente
  • Clientes sin detalle en API → se usa NIT como nombre, no falla
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
from ventas_cliente_report import build_ventas_por_cliente, fetch_invoices
from comisiones_report import build_commission_report
from cartera_report import fetch_cost_centers, fetch_cartera_invoices, build_cartera_rows
from supabase_writer import (
    upsert_ventas,
    upsert_cartera,
    upsert_comisiones_detalle,
    upsert_comisiones_pendientes,
)

_KNOWN_LIMITS = [
    "Documentos DF (Débito Facturación) no disponibles en la API pública de Siigo. "
    "Clientes con este tipo de documento (ej. QUICENO) mostrarán diferencia vs el export de Siigo.",
    "Facturas emitidas el mismo día que el export pueden diferir por desfase de horario. "
    "Se recomienda usar --mes-pasado para períodos cerrados.",
    "Facturas en USD se convierten usando la tasa almacenada en la factura al momento de emisión.",
]


# ── Cálculo de períodos ───────────────────────────────────────────────────────

def _months_in_range(from_date: date, to_date: date) -> list[tuple[date, date]]:
    """Divide un rango en sub-períodos mensuales para evitar límites de la API."""
    months = []
    cursor = from_date.replace(day=1)
    while cursor <= to_date:
        first = max(cursor, from_date)
        # último día del mes
        if cursor.month == 12:
            last = cursor.replace(year=cursor.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            last = cursor.replace(month=cursor.month + 1, day=1) - timedelta(days=1)
        last = min(last, to_date)
        months.append((first, last))
        cursor = last.replace(day=1)
        if cursor.month == 12:
            cursor = cursor.replace(year=cursor.year + 1, month=1, day=1)
        else:
            cursor = cursor.replace(month=cursor.month + 1, day=1)
    return months


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


def _period_tag(from_date: date, to_date: date) -> str:
    if from_date.year == to_date.year and from_date.month == to_date.month:
        return from_date.strftime("%Y%m")
    return f"{from_date.strftime('%Y%m')}_{to_date.strftime('%Y%m')}"


# ── Informes individuales ─────────────────────────────────────────────────────

def _run_ventas_por_cliente(
    config: Any,
    from_date: date,
    to_date: date,
    timeout: int,
) -> dict[str, Any]:
    print(f"\n[Ventas por cliente]  {from_date} → {to_date}")

    print("  Descargando facturas…")
    invoices = fetch_invoices(config, from_date, to_date, timeout=timeout)
    print(f"  {len(invoices)} facturas.")

    print("  Descargando notas crédito…")
    credit_notes = fetch_credit_notes(config, from_date, to_date, timeout=timeout)
    print(f"  {len(credit_notes)} notas crédito.")

    customer_hints: dict[str, dict[str, str]] = {}
    for doc in invoices + credit_notes:
        cust = doc.get("customer")
        if not isinstance(cust, dict):
            continue
        cid = safe_text(cust.get("id"))
        if cid:
            customer_hints[cid] = {"identification": safe_text(cust.get("identification"))}

    print(f"  Descargando detalles de {len(customer_hints)} clientes…")
    customers_map, failures = fetch_customer_details(config, customer_hints, timeout=timeout)
    if failures:
        print(f"  ⚠  {len(failures)} cliente(s) sin detalle — se usará NIT como nombre.", file=sys.stderr)

    rows = build_ventas_por_cliente(invoices, credit_notes, customers_map)
    print(f"  {len(rows)} clientes. Subiendo a Supabase…")

    n = upsert_ventas(rows, from_date, to_date)
    print(f"  ✓ {n} filas en ventas_por_cliente.")

    return {
        "informe": "ventas_por_cliente",
        "estado": "ok",
        "filas": n,
        "clientes": len(rows),
        "facturas": len(invoices),
        "notas_credito": len(credit_notes),
        "clientes_sin_detalle": len(failures),
    }


def _run_comisiones(
    config: Any,
    from_date: date,
    to_date: date,
    timeout: int,
) -> dict[str, Any]:
    print(f"\n[Comisiones]  {from_date} → {to_date}")
    report = build_commission_report(config, from_date, to_date, timeout=timeout)

    today = date.today()
    print("  Subiendo comisiones cobradas a Supabase…")
    n_det = upsert_comisiones_detalle(report["detail"], from_date, to_date)
    print(f"  ✓ {n_det} filas en comisiones_detalle.")

    print("  Subiendo comisiones pendientes a Supabase…")
    n_pte = upsert_comisiones_pendientes(report["pending"], today)
    print(f"  ✓ {n_pte} filas en comisiones_pendientes.")

    meta = report["meta"]
    return {
        "informe": "comisiones",
        "estado": "ok",
        "filas_detalle": n_det,
        "filas_pendientes": n_pte,
        "facturas": meta["invoices"],
        "facturas_pagadas": meta["invoices_pagadas"],
        "facturas_pendientes": meta["invoices_pendientes"],
        "notas_credito": meta["credit_notes"],
        "vendedores": meta["sellers"],
        "comision_pagada": meta["comision_pagada_total"],
        "comision_pendiente": meta["comision_pte_total"],
    }


def _run_cartera(
    config: Any,
    timeout: int,
    cartera_desde: date | None = None,
) -> dict[str, Any]:
    as_of     = date.today()
    from_date = cartera_desde or (as_of - timedelta(days=365 * 3))

    print(f"\n[Cartera]  facturas desde {from_date}  (corte al {as_of})")

    print("  Descargando centros de costo…")
    cost_centers_map = fetch_cost_centers(config, timeout=timeout)
    print(f"  {len(cost_centers_map)} centros de costo.")

    print("  Descargando vendedores…")
    users_map, _ = fetch_users_map(config, timeout=timeout)

    print("  Descargando facturas con saldo pendiente…")
    invoices = fetch_cartera_invoices(config, from_date, as_of, timeout=timeout)
    print(f"  {len(invoices)} facturas con balance > 0.")

    customer_hints: dict[str, dict[str, str]] = {}
    for inv in invoices:
        cust = inv.get("customer")
        if not isinstance(cust, dict):
            continue
        cid = safe_text(cust.get("id"))
        if cid:
            customer_hints[cid] = {"identification": safe_text(cust.get("identification"))}

    print(f"  Descargando detalles de {len(customer_hints)} clientes (para ciudad)…")
    customers_map, failures = fetch_customer_details(config, customer_hints, timeout=timeout)
    if failures:
        print(f"  ⚠  {len(failures)} clientes sin detalle.", file=sys.stderr)

    rows = build_cartera_rows(invoices, cost_centers_map, users_map, customers_map, as_of)
    print(f"  {len(rows)} documentos. Subiendo a Supabase…")

    n = upsert_cartera(rows, as_of)
    print(f"  ✓ {n} filas en cartera.")

    total      = sum(r["Total cartera"] for r in rows)
    vencido    = sum(
        r["Vencido 1 a 30"] + r["Vencido 31 a 60"] +
        r["Vencido 61 a 90"] + r["Vencido más de 91"]
        for r in rows
    )
    por_vencer = sum(r["Saldo por vencer"] for r in rows)

    return {
        "informe": "cartera",
        "estado": "ok",
        "filas": n,
        "corte": as_of.isoformat(),
        "documentos": len(rows),
        "total_cartera": round(total, 2),
        "saldo_vencido": round(vencido, 2),
        "saldo_por_vencer": round(por_vencer, 2),
    }


# ── CLI ───────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Orquestador de informes Siigo → Supabase.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python3 run_informes.py                              # mes actual\n"
            "  python3 run_informes.py --mes-pasado                 # mes anterior\n"
            "  python3 run_informes.py --from-date 2026-01-01 --to-date 2026-06-30\n"
            "  python3 run_informes.py --mes-pasado --solo ventas\n"
        ),
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--mes-actual",  action="store_true",
                      help="Mes en curso (día 1 → ayer) [default]")
    mode.add_argument("--mes-pasado",  action="store_true",
                      help="Mes calendario anterior completo")
    p.add_argument("--from-date", metavar="YYYY-MM-DD")
    p.add_argument("--to-date",   metavar="YYYY-MM-DD")
    p.add_argument("--solo", choices=["ventas", "comisiones", "cartera"],
                   help="Ejecuta solo uno de los informes")
    p.add_argument("--timeout", type=int, default=90,
                   help="Timeout HTTP por página en segundos (default: 90)")
    p.add_argument("--cartera-desde", metavar="YYYY-MM-DD", default=None,
                   help="Fecha inicial para buscar facturas en cartera (default: hace 36 meses)")
    return p.parse_args()


def main() -> int:
    args = _parse_args()

    # ── Resolver período ──────────────────────────────────────────────────────
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
    print(f"  INFORMES SIIGO → SUPABASE  |  {from_date} → {to_date}  |  {tag}")
    print(f"{'='*62}")

    # ── Config ────────────────────────────────────────────────────────────────
    try:
        config = build_siigo_config()
    except (SiigoApiError, RuntimeError) as exc:
        print(f"\nERROR de configuración: {exc}", file=sys.stderr)
        print(
            "Verifica que SIIGO_USERNAME, SIIGO_ACCESS_KEY, SIIGO_PARTNER_ID "
            "y SUPABASE_DB_URL están definidas.",
            file=sys.stderr,
        )
        return 2

    results: list[dict[str, Any]] = []

    # Cuando el rango cubre varios meses, iterar mes a mes para respetar
    # los límites de la API de Siigo y tener datos por mes en Supabase.
    months = _months_in_range(from_date, to_date)
    multi_mes = len(months) > 1
    if multi_mes:
        print(f"  Rango multi-mes: se procesarán {len(months)} meses individualmente.")

    # ── Ventas por cliente ────────────────────────────────────────────────────
    if args.solo in (None, "ventas"):
        for m_from, m_to in months:
            try:
                results.append(_run_ventas_por_cliente(config, m_from, m_to, args.timeout))
            except Exception as exc:
                print(f"\n  ✗ Ventas {m_from}→{m_to} FALLÓ: {exc}", file=sys.stderr)
                results.append({
                    "informe": "ventas_por_cliente",
                    "estado": "error",
                    "periodo": f"{m_from}→{m_to}",
                    "error": str(exc),
                    "traceback": traceback.format_exc(),
                })

    # ── Comisiones ────────────────────────────────────────────────────────────
    if args.solo in (None, "comisiones"):
        for m_from, m_to in months:
            try:
                results.append(_run_comisiones(config, m_from, m_to, args.timeout))
            except Exception as exc:
                print(f"\n  ✗ Comisiones {m_from}→{m_to} FALLÓ: {exc}", file=sys.stderr)
                results.append({
                    "informe": "comisiones",
                    "estado": "error",
                    "periodo": f"{m_from}→{m_to}",
                    "error": str(exc),
                    "traceback": traceback.format_exc(),
                })

    # ── Cartera ───────────────────────────────────────────────────────────────
    if args.solo in (None, "cartera"):
        cartera_desde = (
            date.fromisoformat(args.cartera_desde) if args.cartera_desde else None
        )
        try:
            results.append(_run_cartera(config, args.timeout, cartera_desde))
        except Exception as exc:
            print(f"\n  ✗ Cartera FALLÓ: {exc}", file=sys.stderr)
            results.append({
                "informe": "cartera",
                "estado": "error",
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # ── Resumen ───────────────────────────────────────────────────────────────
    ok_count  = sum(1 for r in results if r["estado"] == "ok")
    err_count = sum(1 for r in results if r["estado"] == "error")

    print(f"\n{'─'*62}")
    for r in results:
        if r["estado"] == "ok":
            filas = r.get("filas") or r.get("filas_detalle", "?")
            print(f"  ✓  {r['informe']:<30} {filas} filas")
        else:
            print(f"  ✗  {r['informe']:<30} ERROR: {r['error']}", file=sys.stderr)

    print(f"\n  {ok_count}/{len(results)} informes subidos a Supabase")
    print(f"  Limitaciones conocidas: {len(_KNOWN_LIMITS)} (ver documentación)")
    if err_count:
        print(
            f"\n  ⚠  {err_count} informe(s) con error. "
            "Re-ejecuta con --solo <nombre> para reintentar.",
            file=sys.stderr,
        )
    print(f"{'─'*62}\n")

    # Imprimir resumen JSON para el log de GitHub Actions
    print(json.dumps({
        "generado_en": datetime.now(timezone.utc).isoformat(),
        "periodo": {"desde": from_date.isoformat(), "hasta": to_date.isoformat()},
        "informes": results,
    }, ensure_ascii=False, indent=2))

    return 0 if err_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
