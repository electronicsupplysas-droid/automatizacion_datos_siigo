# Biblioteca de datos — Siigo Nube API

Referencia completa de los tipos de información disponibles en la API de Siigo y las mejoras propuestas para la integración con Supabase.

---

## Arquitectura objetivo

```
Siigo API  →  Python (extractor)  →  Supabase (almacén)  →  Looker Studio / dashboard
                                          ↑
                               pg_cron (job nativo de Supabase)
                               reemplaza GitHub Actions como scheduler
```

Los **Supabase Edge Functions + pg_cron** permiten disparar la sincronización directamente desde la plataforma, sin depender de un runner externo.

---

## Módulos ya implementados

### 1. Facturas de venta — `v1/invoices`

Endpoint con paginación. Filtros por rango de fechas (`date_start`, `date_end`).

**Campos extraídos actualmente**

| Campo | Descripción |
|---|---|
| `id` | UUID interno de Siigo |
| `name` | Nombre legible de la factura (ej. FV-0012) |
| `prefix` / `number` | Prefijo y consecutivo |
| `date` | Fecha de emisión |
| `customer.id` / `customer.identification` | ID y NIT/CC del cliente |
| `customer.branch_office` | Sucursal del cliente |
| `seller` | ID del vendedor asignado |
| `total` | Valor total de la factura |
| `balance` | Saldo pendiente por pagar |
| `annulled` | Si la factura fue anulada |
| `currency.code` | Código de moneda (COP, USD…) |
| `document.id` | Tipo de documento contable |

**Reportes generados desde este módulo**

| Reporte | Descripción | Salidas |
|---|---|---|
| `sales-report` | Ventas por cliente y mes | JSON · CSV por cliente · CSV por cliente-mes · CSV detalle |
| `billing-report` | Facturación con vendedor | JSON · XLSX · CSV por cliente · CSV por vendedor · CSV cliente-vendedor · CSV por día |
| `cartera-report` | Cuentas por cobrar con aging | JSON · CSV por cliente · CSV detalle de facturas |

---

### 2. Clientes — `v1/customers`

Catálogo maestro de clientes. Se consulta individualmente por UUID para enriquecer facturas.

**Campos disponibles**

| Campo | Descripción |
|---|---|
| `id` | UUID interno |
| `identification` | NIT o cédula |
| `commercial_name` | Nombre comercial |
| `name` | Nombre(s) y apellido(s) |
| `person_type` | Persona Natural / Jurídica |
| `vat_responsible` | Responsable de IVA |
| `email` / `phone` | Contacto |
| `address` | Dirección (línea, ciudad, depto, país) |
| `active` | Estado activo/inactivo |

---

### 3. Usuarios / Vendedores — `v1/users`

Lista completa de usuarios de la empresa. Se usa para resolver el nombre del vendedor en las facturas.

**Campos disponibles**

| Campo | Descripción |
|---|---|
| `id` | UUID interno |
| `first_name` / `last_name` | Nombre completo |
| `username` / `email` | Acceso a Siigo |
| `identification` | Documento de identidad |
| `active` | Estado activo |

---

### 4. Productos — `v1/products`

Catálogo de productos y servicios.

**Campos disponibles**

| Campo | Descripción |
|---|---|
| `id` | UUID interno |
| `code` | Código del producto |
| `name` | Nombre |
| `account_group` | Grupo contable |
| `type` | Producto / Servicio |
| `active` | Estado activo |
| `prices` | Lista de precios por lista de precio |
| `tax_classification` | Clasificación tributaria |

---

## Módulos disponibles — pendientes de implementar

### 5. Líneas de factura (items) — `v1/invoices/{id}/items` o payload completo

Cada factura tiene un arreglo `items` con el detalle de productos vendidos. **Este campo está en el payload raw pero no se extrae hoy.**

**Campos disponibles en `items[]`**

| Campo | Descripción |
|---|---|
| `product.id` | UUID del producto |
| `product.code` | Código del producto |
| `product.name` | Nombre del producto |
| `quantity` | Cantidad facturada |
| `unit_price` | Precio unitario antes de impuestos |
| `discount` | Porcentaje de descuento aplicado |
| `total` | Total de la línea |
| `taxes` | Impuestos aplicados (IVA, INC…) |

**Mejora propuesta**: tabla `siigo_invoice_items` con FK a `siigo_invoices`.  
Habilita: ranking de productos más vendidos, análisis de márgenes, descuentos por cliente.

---

### 6. Notas crédito — `v1/credit-notes`

Devoluciones y descuentos emitidos sobre facturas de venta.

**Campos clave**

| Campo | Descripción |
|---|---|
| `id` | UUID |
| `date` | Fecha de emisión |
| `customer` | Cliente relacionado |
| `invoice` | Factura original (si aplica) |
| `total` | Valor total de la nota |
| `items` | Líneas de devolución/descuento |

**Mejora propuesta**: tabla `siigo_credit_notes`.  
Habilita: ventas netas reales (facturado − devoluciones), tasa de devolución por cliente o producto.

---

### 7. Facturas de compra — `v1/purchase-invoices`

Compras a proveedores. Clave para calcular **margen bruto**.

**Campos clave**

| Campo | Descripción |
|---|---|
| `id` | UUID |
| `date` | Fecha |
| `supplier` | Proveedor |
| `items` | Productos y costos |
| `total` | Total compra |

**Mejora propuesta**: tabla `siigo_purchase_invoices` + tabla `siigo_suppliers`.  
Habilita: margen bruto por producto, costo de ventas, rotación de inventario.

---

### 8. Recibos de caja / Pagos — `v1/payments`

Registros de pagos recibidos de clientes.

**Campos clave**

| Campo | Descripción |
|---|---|
| `id` | UUID |
| `date` | Fecha del pago |
| `customer` | Cliente que pagó |
| `total` | Monto recibido |
| `applied_to` | Facturas a las que se aplicó el pago |

**Mejora propuesta**: tabla `siigo_payments`.  
Habilita: conciliación de cartera en tiempo real, días de pago promedio (DSO), flujo de caja.

---

### 9. Centros de costo — `v1/cost-centers`

Catálogo de centros de costo definidos en Siigo.

**Campos clave**

| Campo | Descripción |
|---|---|
| `id` | UUID |
| `code` | Código contable |
| `name` | Nombre (ej. "Zona Norte", "Proyecto X") |
| `active` | Estado |

**Mejora propuesta**: dimensión `siigo_cost_centers` y FK en facturas.  
Habilita: segmentación de ventas por zona, proyecto o área.

---

### 10. Tipos de documento — `v1/document-types`

Catálogo de tipos de documento contable (FV, NC, RC, etc.).

**Mejora propuesta**: tabla de referencia `siigo_document_types`.  
Habilita: filtrar dashboards por tipo de transacción de forma legible.

---

### 11. Impuestos — `v1/taxes`

Catálogo de impuestos configurados (IVA 19%, IVA 5%, INC, etc.).

**Mejora propuesta**: tabla de referencia `siigo_taxes`.  
Habilita: análisis de impuestos recaudados, declaraciones.

---

### 12. Grupos de cuentas — `v1/account-groups`

Grupos de cuentas contables (inventario, servicios, etc.).

**Mejora propuesta**: dimensión para categorizar productos en reportes.

---

## Propuestas de mejora al esquema actual

### A. Agregar dimensión de líneas de factura

```sql
create table if not exists public.siigo_invoice_items (
    id            bigint generated by default as identity primary key,
    invoice_id    text references public.siigo_invoices(invoice_id),
    line_index    integer not null,
    product_id    text,
    product_code  text,
    product_name  text,
    quantity      numeric(18, 4),
    unit_price    numeric(18, 2),
    discount_pct  numeric(5, 2),
    line_total    numeric(18, 2),
    tax_amount    numeric(18, 2),
    updated_at    timestamptz not null default now()
);
```

### B. Enriquecer clientes con ciudad y departamento

Los campos `city_name` y `state_name` ya están en la tabla `siigo_customers` pero no se usan en las vistas.

```sql
-- Vista de facturación por región
create or replace view public.rpt_siigo_billing_by_region as
select
    i.invoice_month,
    c.state_name  as departamento,
    c.city_name   as ciudad,
    count(*)::integer as invoice_count,
    sum(i.total_amount)::numeric(18,2) as total_sales
from public.siigo_invoices i
left join public.siigo_customers c on c.customer_id = i.customer_id
where not i.annulled
group by i.invoice_month, c.state_name, c.city_name;
```

### C. Agregar meta de ventas por vendedor

Tabla nueva para registrar metas y calcular % de cumplimiento:

```sql
create table if not exists public.siigo_seller_targets (
    id          bigint generated by default as identity primary key,
    seller_id   text references public.siigo_sellers(seller_id),
    period      date not null,        -- primer día del mes
    target_amount numeric(18,2) not null,
    notes       text,
    created_at  timestamptz not null default now()
);
```

```sql
-- Vista de cumplimiento de meta
create or replace view public.rpt_siigo_seller_vs_target as
select
    s.invoice_month,
    s.seller_id,
    s.seller_name,
    s.total_sales as actual_sales,
    t.target_amount,
    round(s.total_sales / nullif(t.target_amount, 0) * 100, 1) as pct_cumplimiento
from public.rpt_siigo_billing_by_seller s
left join public.siigo_seller_targets t
    on t.seller_id = s.seller_id
    and t.period = s.invoice_month;
```

### D. Dimensión de categoría de producto

Añadir columna `product_category` a `siigo_invoice_items` (se puede mapear desde `account_group` del catálogo de productos).

---

## Supabase pg_cron — reemplazo de GitHub Actions

En lugar de usar GitHub Actions como scheduler externo, Supabase ofrece **pg_cron** de forma nativa para invocar Edge Functions o ejecutar SQL periódicamente.

### Activar pg_cron en Supabase

```sql
-- En el SQL Editor de Supabase (requiere rol postgres)
create extension if not exists pg_cron;
```

### Opción 1 — pg_cron llama a una Edge Function

```sql
select cron.schedule(
    'siigo-daily-sync',
    '0 3 * * *',   -- 3 AM Colombia (UTC-5) = 8 AM UTC
    $$
    select
        net.http_post(
            url := 'https://<project>.functions.supabase.co/siigo-sync',
            headers := jsonb_build_object(
                'Authorization', 'Bearer ' || current_setting('app.service_role_key'),
                'Content-Type', 'application/json'
            ),
            body := '{"mode":"incremental"}'::jsonb
        )
    $$
);
```

### Opción 2 — pg_cron ejecuta la sincronización directamente en SQL

Para los datos que ya están en Supabase (ej. refrescar vistas materializadas o mover datos entre tablas staging → producción):

```sql
select cron.schedule(
    'refresh-reports',
    '10 3 * * *',
    'refresh materialized view concurrently public.rpt_siigo_billing_summary'
);
```

### Administrar jobs

```sql
-- Ver todos los jobs
select * from cron.job;

-- Ver historial de ejecuciones
select * from cron.job_run_details order by start_time desc limit 20;

-- Eliminar un job
select cron.unschedule('siigo-daily-sync');
```

---

## Resumen de coberturas de la API

| Módulo | Estado | Prioridad |
|---|---|---|
| Facturas de venta (`v1/invoices`) | Implementado | — |
| Clientes (`v1/customers`) | Implementado | — |
| Vendedores (`v1/users`) | Implementado | — |
| Productos (`v1/products`) | Lectura parcial | Media |
| Líneas de factura (items) | Pendiente | **Alta** |
| Notas crédito (`v1/credit-notes`) | Pendiente | Alta |
| Recibos de caja (`v1/payments`) | Pendiente | Alta |
| Facturas de compra | Pendiente | Media |
| Centros de costo | Pendiente | Media |
| Tipos de documento | Pendiente | Baja |
| Impuestos | Pendiente | Baja |

---

## Dashboards posibles con los datos actuales

| Dashboard | Fuente |
|---|---|
| Ventas mensuales por cliente | `rpt_siigo_billing_by_customer` |
| Ranking de vendedores | `rpt_siigo_billing_by_seller` |
| Cruce cliente-vendedor | `rpt_siigo_billing_by_customer_seller` |
| Tendencia diaria de facturación | `rpt_siigo_billing_by_day` |
| Cartera por antigüedad | `rpt_siigo_cartera` |
| Salud del sync | `rpt_siigo_sync_health` |

## Dashboards adicionales con mejoras propuestas

| Dashboard | Requiere |
|---|---|
| Top productos más vendidos | Líneas de factura |
| Ventas netas (facturas − NC) | Notas crédito |
| Margen bruto por producto | Compras + líneas de factura |
| Mapa de calor por ciudad | Campo `city_name` en customers (ya disponible) |
| Cumplimiento de meta por vendedor | Tabla `siigo_seller_targets` |
| Flujo de caja real | Recibos de caja |
| DSO (días de cobranza promedio) | Recibos de caja + facturas |
