-- schema.sql
-- Ejecutar una sola vez en Supabase → SQL Editor
-- Supabase > Project > SQL Editor > New query → pegar y ejecutar

-- ── Ventas por cliente ────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS ventas_por_cliente (
    id                  BIGSERIAL PRIMARY KEY,
    periodo             DATE        NOT NULL,  -- primer día del período (from_date)
    periodo_fin         DATE        NOT NULL,  -- último día del período (to_date)
    nit                 TEXT,
    cliente             TEXT        NOT NULL,
    num_comprobantes    INTEGER     NOT NULL DEFAULT 0,
    valor_bruto         NUMERIC(18,2) NOT NULL DEFAULT 0,
    descuentos          NUMERIC(18,2) NOT NULL DEFAULT 0,
    subtotal            NUMERIC(18,2) NOT NULL DEFAULT 0,
    impuesto_cargo      NUMERIC(18,2) NOT NULL DEFAULT 0,
    impuesto_retencion  NUMERIC(18,2) NOT NULL DEFAULT 0,
    total               NUMERIC(18,2) NOT NULL DEFAULT 0,
    generado_en         TIMESTAMPTZ NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_ventas_periodo ON ventas_por_cliente (periodo, periodo_fin);
CREATE INDEX IF NOT EXISTS idx_ventas_nit     ON ventas_por_cliente (nit);


-- ── Cartera ───────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS cartera (
    id                  BIGSERIAL PRIMARY KEY,
    snapshot_date       DATE        NOT NULL,  -- fecha de corte (hoy al ejecutar)
    nit                 TEXT,
    cliente             TEXT        NOT NULL,
    sucursal            TEXT,
    documento           TEXT        NOT NULL,  -- nombre de la factura, ej. FV-1-1234
    fecha_vencimiento   DATE,
    centro_costo        TEXT,
    vendedor            TEXT,
    ciudad              TEXT,
    vencido_1_30        NUMERIC(18,2) NOT NULL DEFAULT 0,
    vencido_31_60       NUMERIC(18,2) NOT NULL DEFAULT 0,
    vencido_61_90       NUMERIC(18,2) NOT NULL DEFAULT 0,
    vencido_mas_91      NUMERIC(18,2) NOT NULL DEFAULT 0,
    saldo_por_vencer    NUMERIC(18,2) NOT NULL DEFAULT 0,
    saldo_a_favor       NUMERIC(18,2) NOT NULL DEFAULT 0,
    total_cartera       NUMERIC(18,2) NOT NULL DEFAULT 0,
    generado_en         TIMESTAMPTZ NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_cartera_snapshot  ON cartera (snapshot_date);
CREATE INDEX IF NOT EXISTS idx_cartera_nit        ON cartera (nit);
CREATE INDEX IF NOT EXISTS idx_cartera_documento  ON cartera (documento);


-- ── Comisiones pagadas ────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS comisiones_detalle (
    id              BIGSERIAL PRIMARY KEY,
    periodo         DATE        NOT NULL,  -- from_date del informe
    periodo_fin     DATE        NOT NULL,  -- to_date del informe
    mes_cobro       TEXT        NOT NULL,  -- "2026-05" (mes en que se cobró)
    fecha_factura   DATE,
    factura         TEXT        NOT NULL,  -- nombre, ej. FV-1-1234
    moneda          TEXT        NOT NULL DEFAULT 'COP',
    fecha_pago      DATE,
    recibo          TEXT,                  -- nombre del recibo de caja
    cliente         TEXT,
    vendedor        TEXT,
    vendedor_id     TEXT,
    total_factura   NUMERIC(18,2) NOT NULL DEFAULT 0,
    nc_ids          TEXT,                  -- IDs de notas crédito separados por coma
    total_nc        NUMERIC(18,2) NOT NULL DEFAULT 0,
    saldo_neto      NUMERIC(18,2) NOT NULL DEFAULT 0,
    comision        NUMERIC(18,2) NOT NULL DEFAULT 0,
    generado_en     TIMESTAMPTZ NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_com_det_periodo  ON comisiones_detalle (periodo, periodo_fin);
CREATE INDEX IF NOT EXISTS idx_com_det_vendedor ON comisiones_detalle (vendedor);
CREATE INDEX IF NOT EXISTS idx_com_det_factura  ON comisiones_detalle (factura);


-- ── Comisiones pendientes ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS comisiones_pendientes (
    id                  BIGSERIAL PRIMARY KEY,
    snapshot_date       DATE        NOT NULL,  -- fecha en que se tomó el snapshot
    fecha_factura       DATE,
    factura             TEXT        NOT NULL,
    moneda              TEXT        NOT NULL DEFAULT 'COP',
    cliente             TEXT,
    vendedor            TEXT,
    total_factura       NUMERIC(18,2) NOT NULL DEFAULT 0,
    nc_ids              TEXT,
    total_nc            NUMERIC(18,2) NOT NULL DEFAULT 0,
    saldo_neto          NUMERIC(18,2) NOT NULL DEFAULT 0,
    comision_pendiente  NUMERIC(18,2) NOT NULL DEFAULT 0,
    dias_sin_pago       INTEGER      NOT NULL DEFAULT 0,
    balance_api         NUMERIC(18,2) NOT NULL DEFAULT 0,
    generado_en         TIMESTAMPTZ NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_com_pte_snapshot ON comisiones_pendientes (snapshot_date);
CREATE INDEX IF NOT EXISTS idx_com_pte_factura  ON comisiones_pendientes (factura);
CREATE INDEX IF NOT EXISTS idx_com_pte_vendedor ON comisiones_pendientes (vendedor);
