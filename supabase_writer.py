#!/usr/bin/env python3
"""
Escritura de datos Siigo a Supabase (PostgreSQL).

Tablas objetivo:
  documentos    — una fila por FV o NC, upsert por id_siigo
  recibos_caja  — una fila por línea de pago RC, upsert por (id_recibo, factura_referenciada)

Requiere la variable de entorno:
  SUPABASE_DB_URL  — cadena de conexión completa al connection pooler
                     postgresql://postgres.[ref]:[password]@aws-...:6543/postgres
"""
from __future__ import annotations

import os
from datetime import datetime, timezone
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


# ── documentos ────────────────────────────────────────────────────────────────

def upsert_documentos(rows: list[dict[str, Any]]) -> int:
    """
    Inserta o actualiza documentos por id_siigo.
    En conflicto actualiza balance_cop y ultima_actualizacion para mantener
    el saldo siempre vigente sin perder los datos originales del documento.
    """
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        records = [
            (
                r["id_siigo"],
                r["nombre"],
                r["tipo"],
                r["fecha"],
                r["fecha_vencimiento"],
                r["nit_cliente"],
                r["nombre_cliente"],
                r["ciudad_cliente"],
                r["id_vendedor"],
                r["nombre_vendedor"],
                r["centro_costo"],
                r["moneda"],
                r["tasa_cambio"],
                r["valor_bruto_cop"],
                r["descuentos_cop"],
                r["subtotal_cop"],
                r["impuesto_cop"],
                r["total_cop"],
                r["balance_cop"],
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO documentos
              (id_siigo, nombre, tipo, fecha, fecha_vencimiento,
               nit_cliente, nombre_cliente, ciudad_cliente,
               id_vendedor, nombre_vendedor, centro_costo,
               moneda, tasa_cambio,
               valor_bruto_cop, descuentos_cop, subtotal_cop, impuesto_cop,
               total_cop, balance_cop, ultima_actualizacion)
            VALUES %s
            ON CONFLICT (id_siigo) DO UPDATE SET
              balance_cop          = EXCLUDED.balance_cop,
              nombre_vendedor      = EXCLUDED.nombre_vendedor,
              fecha_vencimiento    = EXCLUDED.fecha_vencimiento,
              ultima_actualizacion = EXCLUDED.ultima_actualizacion
            """,
            records,
        )
    return len(records)


# ── recibos_caja ──────────────────────────────────────────────────────────────

def upsert_recibos(rows: list[dict[str, Any]]) -> int:
    """
    Inserta o actualiza recibos por (id_recibo, factura_referenciada).
    En conflicto actualiza monto_cop y fecha_cobro.
    """
    if not rows:
        return 0
    ts = _ts()
    with _conn() as conn, conn.cursor() as cur:
        records = [
            (
                r["id_recibo"],
                r["fecha_cobro"],
                r["factura_referenciada"],
                r["nit_cliente"],
                r["nombre_cliente"],
                r["id_vendedor"],
                r["nombre_vendedor"],
                r["monto_cop"],
                ts,
            )
            for r in rows
        ]
        psycopg2.extras.execute_values(
            cur,
            """
            INSERT INTO recibos_caja
              (id_recibo, fecha_cobro, factura_referenciada,
               nit_cliente, nombre_cliente,
               id_vendedor, nombre_vendedor,
               monto_cop, ultima_actualizacion)
            VALUES %s
            ON CONFLICT (id_recibo, factura_referenciada) DO UPDATE SET
              fecha_cobro          = EXCLUDED.fecha_cobro,
              monto_cop            = EXCLUDED.monto_cop,
              ultima_actualizacion = EXCLUDED.ultima_actualizacion
            """,
            records,
        )
    return len(records)
