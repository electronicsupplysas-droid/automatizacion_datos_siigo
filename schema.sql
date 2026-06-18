-- schema.sql — Ejecutar en Supabase → SQL Editor
-- Elimina el esquema anterior y crea el nuevo diseño desagregado.
-- Correr una sola vez (o cuando se quiera resetear las tablas).

-- ── Eliminar tablas anteriores (si existen) ───────────────────────────────────
DROP TABLE IF EXISTS ventas_por_cliente    CASCADE;
DROP TABLE IF EXISTS cartera               CASCADE;
DROP TABLE IF EXISTS comisiones_detalle    CASCADE;
DROP TABLE IF EXISTS comisiones_pendientes CASCADE;

-- ── documentos ────────────────────────────────────────────────────────────────
-- Una fila por documento (FV = Factura Venta, NC = Nota Crédito).
-- Upsert por id_siigo: balance_cop se actualiza diariamente vía cartera_update.
--
-- Notas de diseño:
--   total_cop   = doc.total × tasa_cambio  (campo oficial de Siigo)
--   balance_cop = doc.balance × tasa_cambio (solo FV; NC siempre 0)
--   valor_bruto/descuentos/impuesto: calculados desde items (mejor esfuerzo)
--   fecha_vencimiento: payments[].due_date máximo, nunca el campo raíz
--   DF y ND no disponibles vía API → nunca aparecen en esta tabla
CREATE TABLE IF NOT EXISTS documentos (
    id                   BIGSERIAL    PRIMARY KEY,
    id_siigo             TEXT         NOT NULL UNIQUE,
    nombre               TEXT         NOT NULL,             -- "FV-1-1234"
    tipo                 TEXT         NOT NULL CHECK (tipo IN ('FV', 'NC')),
    fecha                DATE,
    fecha_vencimiento    DATE,
    nit_cliente          TEXT,
    nombre_cliente       TEXT,
    ciudad_cliente       TEXT,
    id_vendedor          TEXT,
    nombre_vendedor      TEXT,
    centro_costo         TEXT,
    moneda               TEXT         NOT NULL DEFAULT 'COP',
    tasa_cambio          NUMERIC(12,6) NOT NULL DEFAULT 1,
    valor_bruto_cop      NUMERIC(18,2) NOT NULL DEFAULT 0,
    descuentos_cop       NUMERIC(18,2) NOT NULL DEFAULT 0,
    subtotal_cop         NUMERIC(18,2) NOT NULL DEFAULT 0,
    impuesto_cop         NUMERIC(18,2) NOT NULL DEFAULT 0,
    total_cop            NUMERIC(18,2) NOT NULL DEFAULT 0,
    balance_cop          NUMERIC(18,2) NOT NULL DEFAULT 0,  -- actualizado diariamente
    ultima_actualizacion TIMESTAMPTZ  NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_doc_tipo            ON documentos (tipo);
CREATE INDEX IF NOT EXISTS idx_doc_fecha           ON documentos (fecha);
CREATE INDEX IF NOT EXISTS idx_doc_nit             ON documentos (nit_cliente);
CREATE INDEX IF NOT EXISTS idx_doc_vendedor        ON documentos (nombre_vendedor);
CREATE INDEX IF NOT EXISTS idx_doc_balance         ON documentos (balance_cop) WHERE balance_cop > 0;
CREATE INDEX IF NOT EXISTS idx_doc_fecha_venc      ON documentos (fecha_vencimiento);


-- ── recibos_caja ──────────────────────────────────────────────────────────────
-- Una fila por línea de pago dentro de un voucher (RC = Recibo de Caja).
-- Un mismo RC puede pagar varias facturas → varias filas con el mismo id_recibo.
-- factura_referenciada: "{prefix}-{consecutive}" igual que documentos.nombre
CREATE TABLE IF NOT EXISTS recibos_caja (
    id                   BIGSERIAL    PRIMARY KEY,
    id_recibo            TEXT         NOT NULL,             -- "RC-2-123"
    fecha_cobro          DATE,
    factura_referenciada TEXT         NOT NULL,             -- "FV-1-1234"
    nit_cliente          TEXT,
    nombre_cliente       TEXT,
    id_vendedor          TEXT,
    nombre_vendedor      TEXT,
    monto_cop            NUMERIC(18,2) NOT NULL DEFAULT 0,
    ultima_actualizacion TIMESTAMPTZ  NOT NULL,
    UNIQUE (id_recibo, factura_referenciada)
);

CREATE INDEX IF NOT EXISTS idx_rc_fecha_cobro  ON recibos_caja (fecha_cobro);
CREATE INDEX IF NOT EXISTS idx_rc_factura_ref  ON recibos_caja (factura_referenciada);
CREATE INDEX IF NOT EXISTS idx_rc_nit          ON recibos_caja (nit_cliente);
CREATE INDEX IF NOT EXISTS idx_rc_vendedor     ON recibos_caja (nombre_vendedor);


-- ── Vistas útiles para Looker Studio ─────────────────────────────────────────

-- Ventas por mes y cliente (equivale al antiguo informe ventas_por_cliente)
CREATE OR REPLACE VIEW v_ventas_mes AS
SELECT
    DATE_TRUNC('month', fecha)::DATE AS mes,
    nit_cliente,
    nombre_cliente,
    ciudad_cliente,
    nombre_vendedor,
    centro_costo,
    moneda,
    SUM(CASE WHEN tipo = 'FV' THEN total_cop ELSE 0 END)  AS total_fv_cop,
    SUM(CASE WHEN tipo = 'NC' THEN total_cop ELSE 0 END)  AS total_nc_cop,
    SUM(CASE WHEN tipo = 'FV' THEN total_cop
             WHEN tipo = 'NC' THEN -total_cop END)         AS neto_cop,
    COUNT(CASE WHEN tipo = 'FV' THEN 1 END)               AS num_facturas,
    COUNT(CASE WHEN tipo = 'NC' THEN 1 END)               AS num_nc
FROM documentos
WHERE fecha IS NOT NULL
GROUP BY 1, 2, 3, 4, 5, 6, 7;

-- Cartera vigente (snapshot en tiempo real)
CREATE OR REPLACE VIEW v_cartera AS
SELECT
    d.nombre               AS documento,
    d.fecha                AS fecha_factura,
    d.fecha_vencimiento,
    d.nit_cliente,
    d.nombre_cliente,
    d.ciudad_cliente,
    d.nombre_vendedor,
    d.centro_costo,
    d.moneda,
    d.total_cop,
    d.balance_cop,
    CURRENT_DATE - d.fecha_vencimiento         AS dias_vencido,
    CASE
        WHEN d.fecha_vencimiento IS NULL THEN 'Sin vencimiento'
        WHEN CURRENT_DATE <= d.fecha_vencimiento THEN 'Por vencer'
        WHEN CURRENT_DATE - d.fecha_vencimiento <= 30  THEN 'Vencido 1-30'
        WHEN CURRENT_DATE - d.fecha_vencimiento <= 60  THEN 'Vencido 31-60'
        WHEN CURRENT_DATE - d.fecha_vencimiento <= 90  THEN 'Vencido 61-90'
        ELSE 'Vencido >91'
    END                                        AS bucket_vencimiento,
    d.ultima_actualizacion
FROM documentos d
WHERE d.tipo = 'FV'
  AND d.balance_cop > 0;

-- Comisiones cobradas (FV con recibo de caja asociado)
CREATE OR REPLACE VIEW v_comisiones_cobradas AS
SELECT
    rc.fecha_cobro,
    DATE_TRUNC('month', rc.fecha_cobro)::DATE  AS mes_cobro,
    rc.id_recibo,
    rc.factura_referenciada,
    d.fecha                                    AS fecha_factura,
    d.nit_cliente,
    d.nombre_cliente,
    COALESCE(rc.nombre_vendedor, d.nombre_vendedor) AS vendedor,
    d.moneda,
    d.total_cop                                AS total_factura_cop,
    rc.monto_cop                               AS monto_cobrado_cop,
    ROUND(rc.monto_cop * 0.03, 2)             AS comision_cop
FROM recibos_caja rc
LEFT JOIN documentos d ON d.nombre = rc.factura_referenciada;

-- Comisiones pendientes (FV sin recibo de caja)
CREATE OR REPLACE VIEW v_comisiones_pendientes AS
SELECT
    d.nombre               AS factura,
    d.fecha                AS fecha_factura,
    d.nit_cliente,
    d.nombre_cliente,
    d.nombre_vendedor      AS vendedor,
    d.moneda,
    d.total_cop,
    d.balance_cop,
    ROUND(d.balance_cop * 0.03, 2)            AS comision_pendiente,
    CURRENT_DATE - d.fecha::DATE              AS dias_sin_cobro
FROM documentos d
WHERE d.tipo = 'FV'
  AND d.balance_cop > 0
  AND NOT EXISTS (
      SELECT 1 FROM recibos_caja rc
      WHERE rc.factura_referenciada = d.nombre
  );
