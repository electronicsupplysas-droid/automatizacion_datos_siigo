# Integración de solo lectura con Siigo Nube

Este repositorio quedó en dos capas:

- **Exploración y pruebas controladas** de Siigo en modo solo lectura.
- **Esqueleto de servicio automatizado** para sincronizar datos a Supabase y alimentar dashboards.

El objetivo sigue siendo el mismo: **leer y descargar información**, sin crear documentos por error.

## Arquitectura recomendada

La opción base para cliente real es:

- `Siigo API` como fuente.
- `Python` como extractor.
- `Supabase` como repositorio operativo.
- `GitHub Actions` como scheduler.
- `Looker Studio` sobre vistas `rpt_siigo_*`.

Archivos principales de esta capa:

- [siigo_core.py](/Users/stivenjohanhurtado/Integracion siigo nube/siigo_core.py)
- [sync_siigo_to_supabase.py](/Users/stivenjohanhurtado/Integracion siigo nube/sync_siigo_to_supabase.py)
- [sql/001_base_schema.sql](/Users/stivenjohanhurtado/Integracion siigo nube/sql/001_base_schema.sql)
- [sql/002_reporting_views.sql](/Users/stivenjohanhurtado/Integracion siigo nube/sql/002_reporting_views.sql)
- [.github/workflows/siigo_supabase_sync.yml](/Users/stivenjohanhurtado/Integracion siigo nube/.github/workflows/siigo_supabase_sync.yml)
- [docs/supabase_service.md](/Users/stivenjohanhurtado/Integracion siigo nube/docs/supabase_service.md)

## Sincronización a Supabase

Primera corrida manual por rango:

```bash
python3 sync_siigo_to_supabase.py --from-date 2026-04-01 --to-date 2026-04-30
```

Corrida incremental:

```bash
python3 sync_siigo_to_supabase.py
```

La ventana incremental recomendada es amplia, por ejemplo `90` días, porque el saldo de una factura puede cambiar después de emitida.

## Qué hace

- Obtiene token de autenticación en Siigo.
- Ejecuta únicamente consultas `GET`.
- Bloquea paths que no empiecen por `v1/`.
- No implementa `POST`, `PUT`, `PATCH` ni `DELETE`.

La única petición `POST` que realiza es la de autenticación (`/auth` o `/v1/auth`), necesaria para obtener el token.

## Configuración

1. Crea tu archivo de entorno:

```bash
cp .env.example .env
```

2. Completa estas variables:

- `SIIGO_USERNAME`
- `SIIGO_ACCESS_KEY`
- `SIIGO_PARTNER_ID`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

Si tu Python tiene problemas de certificados SSL, el script intenta usar automáticamente `/etc/ssl/cert.pem` en macOS. Si en otra máquina necesitas forzar un bundle distinto, define:

- `SIIGO_CA_FILE=/ruta/al/cert.pem`

## Pruebas recomendadas

Validar credenciales:

```bash
python3 siigo_readonly_test.py auth-test
```

Consultar primeros clientes:

```bash
python3 siigo_readonly_test.py get v1/customers --param page=1 --param page_size=5
```

Consultar productos y guardar salida:

```bash
python3 siigo_readonly_test.py get v1/products --param page=1 --param page_size=10 --output output/products_page_1.json
```

Exportar todas las páginas de productos en un solo JSON:

```bash
python3 siigo_readonly_test.py export v1/products --param page=1 --param page_size=25 --output output/products_all.json
```

Exportar solo 2 páginas para prueba:

```bash
python3 siigo_readonly_test.py export v1/products --param page=1 --param page_size=25 --max-pages 2 --output output/products_sample.json
```

Construir reporte de ventas por cliente:

```bash
python3 siigo_readonly_test.py sales-report --from-date 2026-01-01 --to-date 2026-06-30 --output output/ventas.json
```

El comando `sales-report` genera:

- Un JSON con `summary`, `by_customer`, `by_customer_month` e `invoices`.
- Un CSV agregado por cliente.
- Un CSV agregado por cliente y mes.
- Un CSV con el detalle de facturas.

Los CSV se exportan con `UTF-8 BOM` y formato regional amigable para Excel en español:

- Separador de columnas: `;`
- Separador decimal: `,`
- Booleanos: `Sí` / `No`

Reporte detallado de facturación:

```bash
python3 siigo_readonly_test.py billing-report --from-date 2026-01-01 --to-date 2026-06-30 --output output/facturacion.json
```

El comando `billing-report` genera:

- Un archivo `.xlsx` con varias hojas y números reales listos para sumar/filtrar.
- Un JSON con `summary`, `by_customer`, `by_seller`, `by_customer_seller`, `by_day` e `invoices`.
- Un CSV agregado por cliente.
- Un CSV agregado por vendedor.
- Un CSV agregado por cliente-vendedor.
- Un CSV agregado por día.
- Un CSV con el detalle de facturas.

Consultar un cliente por id:

```bash
python3 siigo_readonly_test.py get v1/customers/UUID_DEL_CLIENTE
```

Reporte de cartera (cuentas por cobrar):

```bash
python3 siigo_readonly_test.py cartera-report --from-date 2025-01-01 --to-date 2026-06-30 --output output/cartera.json
```

El comando `cartera-report` genera:

- Un JSON con `summary`, `by_bucket`, `by_customer` e `invoices`.
- Un CSV agregado por cliente con saldo total y desglose por rango de vencimiento.
- Un CSV con el detalle de facturas pendientes.

Los rangos de cartera son:

- **Corriente**: 0-30 días
- **31-60 días**: próximo a vencer
- **61-90 días**: vencido
- **Más de 90 días**: crítico

Solo se incluyen facturas no anuladas con saldo pendiente (`balance > 0`). Se recomienda un ventana de 12 meses para capturar facturas antiguas impagas.

Reporte de comisiones por vendedor (fecha real de cobro):

```bash
python3 comisiones_report.py --from-date 2026-01-01 --to-date 2026-06-30 \
    --output output/comisiones.xlsx
```

El comando `comisiones-report` usa los **recibos de caja** (`v1/vouchers`) para
determinar el mes en que se cobró cada factura. La comisión (3%) cae en el mes
del recibo, no en el mes de la factura.

Genera un `.xlsx` con cuatro hojas:

- **Resumen** — comisión total por vendedor en el período
- **Pivot Comisiones** — tabla cruzada vendedor × mes
- **Por Mes** — comisión por vendedor por mes de cobro
- **Detalle Recibos** — cada línea de recibo con su factura, monto y comisión

---

## Ejecución recurrente (orquestador)

El punto de entrada para corridas programadas es `run_informes.py`.
Genera ventas por cliente + comisiones en un solo comando con manejo robusto de errores:

```bash
# Mes en curso (día 1 → ayer) — modo por defecto
python3 run_informes.py

# Mes anterior completo — recomendado para cierre mensual
python3 run_informes.py --mes-pasado

# Rango personalizado
python3 run_informes.py --from-date 2026-01-01 --to-date 2026-06-30

# Re-correr solo un informe fallido
python3 run_informes.py --mes-pasado --solo ventas
python3 run_informes.py --mes-pasado --solo comisiones
```

Salidas:
- `output/informes/ventas_por_cliente_YYYYMM.xlsx` — con fecha en el nombre (archivable)
- `output/informes/comisiones_YYYYMM.xlsx`
- `output/informes/ventas_por_cliente.xlsx` — copia del más reciente
- `output/informes/comisiones.xlsx`
- `output/logs/run_YYYY-MM-DD_HHMMSS.json` — log de cada corrida

Robustez incorporada:
- **Reintentos automáticos**: HTTP 429/5xx y errores de red → backoff 5/15/30 s
- **Aislamiento de fallos**: si un informe falla, el otro igual se genera
- **Bloqueo de concurrencia**: segunda instancia sale con error si hay una activa
- **Archivo en Excel**: error claro si el `.xlsx` está bloqueado por Office
- **Documentos DF**: advertencia en log, no falla (la API de Siigo no los expone)
- **Phantom IVA**: se usa `doc.total` de Siigo en lugar de recalcular

Código de salida: `0` = todo ok · `1` = algún informe falló · `2` = error de config.

---

## Estado actual

Este repo ya tiene:

1. Cliente de solo lectura para Siigo.
2. Reportes locales en JSON, CSV y XLSX.
3. Sincronizador incremental hacia Supabase (facturas + clientes + vendedores + recibos de caja).
4. SQL de tablas y vistas para dashboards (incluyendo cartera y comisiones por fecha real de cobro).
5. Workflow programado con GitHub Actions.
6. Reporte de comisiones basado en recibos de caja (`v1/vouchers`).
