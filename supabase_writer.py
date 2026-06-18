#!/usr/bin/env python3
"""
Escritura de informes Siigo a Supabase (PostgreSQL).

Requiere la variable de entorno:
  SUPABASE_DB_URL  — cadena de conexión completa
                     postgresql://postgres.[ref]:[password]@aws-...:6543/postgres

Las tablas deben existir previamente (ver schema.sql en el repo).
"""
from __future__ import annotations

import os
from datetime import date, datetime, timezone
from typing import Any

import psycopg2
import psycopg2.extras


def _conn() -> "psycopg2.connection":
    url = os.environ.get("SUPABASE_DB_URL")
    if not url:
        raise RuntimeError(
            "La variable de entorno SUPABASE_DB_URL no está definida. "
            "Agrégala en GitHub → Settings → Secrets and variables → Actions."
        )
    return psycopg2.connect(url, sslmode="require")


def _ts() -> str:
    return datetime.now(timezone.utc).isoformat()


# ── Ventas por cliente ─────────────────────────────────────────────────────────

def upsert_ventas(rows: list[dict[str, Any]], from_date: date, to_date: date) -> int:
    """Reemplaza las filas de ventas para el período dado (delete + insert)."""
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        cur.execute(
            "DELETE FROM ventas_por_cliente WHERE periodo = %s AND periodo_fin = %s",
            (from_date, to_date),
        )
        records = [
            (
                from_date,
                to_date,
                r["Identificación"],
                r["Cliente"],
                r["Número de comprobantes"],
                r["Valor bruto"],
                r["Descuentos por item"],
                r["Subtotal"],
                r["Impuesto cargo"],
                r["Impuesto retención"],
                r["Total"],
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO ventas_por_cliente
              (periodo, periodo_fin, nit, cliente, num_comprobantes,
               valor_bruto, descuentos, subtotal, impuesto_cargo,
               impuesto_retencion, total, generado_en)
            VALUES %s
            """,
            records,
        )
    return len(records)


# ── Cartera ────────────────────────────────────────────────────────────────────

def upsert_cartera(rows: list[dict[str, Any]], snapshot_date: date) -> int:
    """Reemplaza el snapshot de cartera para la fecha de corte dada."""
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM cartera WHERE snapshot_date = %s", (snapshot_date,))
        records = [
            (
                snapshot_date,
                r["Identificación"],
                r["Cliente"],
                r.get("Sucursal") or None,
                r["Documento"],
                r.get("Fecha vencimiento") or None,
                r.get("Centro de costo") or None,
                r.get("Vendedor") or None,
                r.get("Ciudad") or None,
                r["Vencido 1 a 30"],
                r["Vencido 31 a 60"],
                r["Vencido 61 a 90"],
                r["Vencido más de 91"],
                r["Saldo por vencer"],
                r["Saldo a favor"],
                r["Total cartera"],
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO cartera
              (snapshot_date, nit, cliente, sucursal, documento,
               fecha_vencimiento, centro_costo, vendedor, ciudad,
               vencido_1_30, vencido_31_60, vencido_61_90, vencido_mas_91,
               saldo_por_vencer, saldo_a_favor, total_cartera, generado_en)
            VALUES %s
            """,
            records,
        )
    return len(records)


# ── Comisiones pagadas ─────────────────────────────────────────────────────────

def upsert_comisiones_detalle(
    rows: list[dict[str, Any]],
    from_date: date,
    to_date: date,
) -> int:
    """Reemplaza el detalle de comisiones cobradas para el período dado."""
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        cur.execute(
            "DELETE FROM comisiones_detalle WHERE periodo = %s AND periodo_fin = %s",
            (from_date, to_date),
        )
        records = [
            (
                from_date,
                to_date,
                r["mes_cobro"],
                r["fecha_factura"],
                r["factura"],
                r.get("moneda") or "COP",
                r["fecha_pago"],
                r["recibo"],
                r["cliente"],
                r.get("vendedor") or "",
                r.get("vendedor_id") or "",
                r["total_factura"],
                r.get("nc_ids") or None,
                r["total_nc"],
                r["saldo_neto"],
                r["comision"],
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO comisiones_detalle
              (periodo, periodo_fin, mes_cobro, fecha_factura, factura, moneda,
               fecha_pago, recibo, cliente, vendedor, vendedor_id,
               total_factura, nc_ids, total_nc, saldo_neto, comision, generado_en)
            VALUES %s
            """,
            records,
        )
    return len(records)


# ── Comisiones pendientes ──────────────────────────────────────────────────────

def upsert_comisiones_pendientes(
    rows: list[dict[str, Any]],
    snapshot_date: date,
) -> int:
    """Reemplaza las comisiones pendientes del snapshot de hoy."""
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        cur.execute(
            "DELETE FROM comisiones_pendientes WHERE snapshot_date = %s",
            (snapshot_date,),
        )
        records = [
            (
                snapshot_date,
                r["fecha_factura"],
                r["factura"],
                r.get("moneda") or "COP",
                r["cliente"],
                r.get("vendedor") or "",
                r["total_factura"],
                r.get("nc_ids") or None,
                r["total_nc"],
                r["saldo_neto"],
                r["comision_pte"],
                r["dias_sin_pago"],
                r.get("balance_api") or 0,
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO comisiones_pendientes
              (snapshot_date, fecha_factura, factura, moneda, cliente, vendedor,
               total_factura, nc_ids, total_nc, saldo_neto, comision_pendiente,
               dias_sin_pago, balance_api, generado_en)
            VALUES %s
            """,
            records,
        )
    return len(records)
