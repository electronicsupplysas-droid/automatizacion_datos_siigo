#!/usr/bin/env python3
"""
Construcción de filas planas desde la API de Siigo para las tablas:
  • documentos   — una fila por FV o NC (facturas y notas crédito)
  • recibos_caja — una fila por línea de pago dentro de un voucher (RC)

Principios de exactitud heredados de iteraciones anteriores:
  - total_cop   = doc.total × tasa_cambio  (campo oficial de Siigo, evita phantom IVA)
  - balance_cop = doc.balance × tasa_cambio (saldo real de la factura)
  - valor_bruto/descuentos/impuesto calculados desde items (mejor esfuerzo)
  - Fecha de vencimiento: payments[].due_date (máximo), nunca el campo raíz (siempre vacío)
  - Documentos anulados excluidos
  - Facturas en USD: tasa de cambio desde invoice_exchange_rate()
  - DF/ND: no disponibles en la API pública — se omiten sin fallar
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

from siigo_core import (
    customer_display_name,
    fetch_paginated_results,
    invoice_exchange_rate,
    safe_text,
    to_decimal,
    user_display_name,
)


def _rfc3339(d: date, end_of_day: bool = False) -> str:
    suffix = "T23:59:59Z" if end_of_day else "T00:00:00Z"
    return f"{d.isoformat()}{suffix}"

MONEY = Decimal("0.01")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _city_from_customer(cust_body: dict | None) -> str:
    if not isinstance(cust_body, dict):
        return ""
    addr = cust_body.get("address")
    if not isinstance(addr, dict):
        return ""
    city = addr.get("city")
    if not isinstance(city, dict):
        return ""
    return safe_text(city.get("city_name")) or ""


def _due_date(doc: dict) -> str | None:
    """Fecha de vencimiento real: máximo de payments[].due_date."""
    payments = doc.get("payments") or []
    dates = [
        p["due_date"][:10]
        for p in payments
        if isinstance(p, dict) and p.get("due_date")
    ]
    return max(dates) if dates else None


def _moneda(doc: dict) -> str:
    info = doc.get("currency")
    code = safe_text(info.get("code")) if isinstance(info, dict) else None
    return code or "COP"


def _items_breakdown(doc: dict, fx: Decimal) -> tuple[float, float, float, float]:
    """Devuelve (valor_bruto, descuentos, subtotal, impuesto) en COP."""
    valor_bruto = Decimal("0")
    descuentos  = Decimal("0")
    impuesto    = Decimal("0")
    for item in doc.get("items") or []:
        if not isinstance(item, dict):
            continue
        qty   = to_decimal(item.get("quantity", 0))
        price = to_decimal(item.get("price", 0))
        disc  = to_decimal(item.get("discount", 0))
        taxes = sum(
            to_decimal(t.get("total", 0))
            for t in (item.get("taxes") or [])
            if isinstance(t, dict)
        )
        valor_bruto += qty * price
        descuentos  += disc
        impuesto    += taxes

    vb_cop  = (valor_bruto * fx).quantize(MONEY, rounding=ROUND_HALF_UP)
    dc_cop  = (descuentos  * fx).quantize(MONEY, rounding=ROUND_HALF_UP)
    sub_cop = (vb_cop - dc_cop)
    imp_cop = (impuesto * fx).quantize(MONEY, rounding=ROUND_HALF_UP)
    return float(vb_cop), float(dc_cop), float(sub_cop), float(imp_cop)


# ── Fetch vouchers (raw) ──────────────────────────────────────────────────────

def fetch_vouchers_raw(
    config: Any,
    from_date: date,
    to_date: date,
    timeout: int = 90,
) -> list[dict]:
    """
    Descarga todos los recibos de caja (RC) entre from_date y hoy.
    La fecha final siempre es hoy para capturar pagos recientes de facturas antiguas.
    """
    fetch_to = date.today()
    fetched  = fetch_paginated_results(
        config,
        "v1/vouchers",
        {
            "page":          "1",
            "page_size":     "100",
            "created_start": _rfc3339(from_date),
            "created_end":   _rfc3339(fetch_to, end_of_day=True),
        },
        timeout=timeout,
    )
    return [v for v in fetched.get("results", []) if isinstance(v, dict)]


# ── Construcción de filas ─────────────────────────────────────────────────────

def build_documentos_rows(
    invoices: list[dict],
    credit_notes: list[dict],
    customers_map: dict[str, dict],
    users_map: dict[str, dict],
    cost_centers_map: dict[int, str],
) -> list[dict[str, Any]]:
    """
    Una fila plana por documento (FV o NC).
    total_cop usa doc.total (oficial) para evitar phantom IVA.
    balance_cop solo aplica a FV; para NC queda en 0.
    """
    rows: list[dict[str, Any]] = []

    for tipo, docs in [("FV", invoices), ("NC", credit_notes)]:
        for doc in docs:
            if doc.get("annulled"):
                continue

            fx = invoice_exchange_rate(doc)

            # ── Cliente ───────────────────────────────────────────────────────
            cust      = doc.get("customer") or {}
            cid       = safe_text(cust.get("id"))
            nit       = safe_text(cust.get("identification"))
            cust_body = customers_map.get(cid)
            cname     = customer_display_name(cust_body, nit or cid)
            city      = _city_from_customer(cust_body)

            # ── Vendedor ──────────────────────────────────────────────────────
            seller_id   = safe_text(doc.get("seller"))
            seller_body = users_map.get(seller_id) if seller_id else None
            seller_name = user_display_name(seller_body, "") if seller_body else ""

            # ── Centro de costo ───────────────────────────────────────────────
            cc_id   = doc.get("cost_center")
            cc_name = cost_centers_map.get(cc_id, "") if cc_id is not None else ""

            # ── Montos ────────────────────────────────────────────────────────
            # total_cop: campo oficial de Siigo (evita discrepancias por phantom IVA)
            total_cop   = float((to_decimal(doc.get("total", 0)) * fx).quantize(MONEY, rounding=ROUND_HALF_UP))
            # balance_cop: saldo pendiente real (solo FV lo expone)
            balance_cop = float((to_decimal(doc.get("balance", 0)) * fx).quantize(MONEY, rounding=ROUND_HALF_UP))
            vb, dc, sub, imp = _items_breakdown(doc, fx)

            rows.append({
                "id_siigo":          safe_text(doc.get("id")),
                "nombre":            safe_text(doc.get("name")),
                "tipo":              tipo,
                "fecha":             (safe_text(doc.get("date")) or "")[:10] or None,
                "fecha_vencimiento": _due_date(doc),
                "nit_cliente":       nit or None,
                "nombre_cliente":    cname or None,
                "ciudad_cliente":    city or None,
                "id_vendedor":       seller_id or None,
                "nombre_vendedor":   seller_name or None,
                "centro_costo":      cc_name or None,
                "moneda":            _moneda(doc),
                "tasa_cambio":       float(fx),
                "valor_bruto_cop":   vb,
                "descuentos_cop":    dc,
                "subtotal_cop":      sub,
                "impuesto_cop":      imp,
                "total_cop":         total_cop,
                "balance_cop":       balance_cop,
            })

    return rows


def build_recibos_rows(
    vouchers: list[dict],
    users_map: dict[str, dict],
    customers_map: dict[str, dict],
) -> list[dict[str, Any]]:
    """
    Una fila por línea de pago dentro de un voucher (RC).
    Un mismo RC puede pagar varias facturas → varias filas con el mismo id_recibo.
    factura_referenciada: "{prefix}-{consecutive}" igual que el campo name de la FV.
    """
    rows: list[dict[str, Any]] = []

    for v in vouchers:
        rc_name = safe_text(v.get("name")) or ""
        rc_date = (safe_text(v.get("date")) or "")[:10] or None

        # Vendedor del recibo
        seller_id   = safe_text(v.get("seller"))
        seller_body = users_map.get(seller_id) if seller_id else None
        seller_name = user_display_name(seller_body, "") if seller_body else ""

        # Cliente del recibo
        cust      = v.get("customer") or {}
        cid       = safe_text(cust.get("id"))
        nit       = safe_text(cust.get("identification"))
        cust_body = customers_map.get(cid)
        cname     = customer_display_name(cust_body, nit or cid)

        for item in v.get("items") or []:
            if not isinstance(item, dict):
                continue
            due = item.get("due")
            if not isinstance(due, dict):
                continue

            prefix      = safe_text(due.get("prefix")) or ""
            consecutive = due.get("consecutive")
            if not prefix or consecutive is None:
                continue

            factura_ref = f"{prefix}-{consecutive}"
            monto       = float(to_decimal(item.get("value", 0)).quantize(MONEY, rounding=ROUND_HALF_UP))

            rows.append({
                "id_recibo":            rc_name,
                "fecha_cobro":          rc_date,
                "factura_referenciada": factura_ref,
                "nit_cliente":          nit or None,
                "nombre_cliente":       cname or None,
                "id_vendedor":          seller_id or None,
                "nombre_vendedor":      seller_name or None,
                "monto_cop":            monto,
            })

    return rows
