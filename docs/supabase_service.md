# Servicio recomendado

Arquitectura propuesta:

- `Siigo API` como fuente de lectura.
- `sync_siigo_to_supabase.py` como extractor incremental.
- `Supabase Postgres` como repositorio maestro.
- `GitHub Actions` como scheduler diario.
- `Looker Studio` conectado a las vistas `rpt_siigo_*`.

## Por qué así

- Mantienes el control del código en Python, que ya validamos contra Siigo.
- Evitas usar Google Sheets como base de datos.
- Tienes historial, trazabilidad y reintentos simples.
- Los dashboards consumen tablas/vistas estables, no archivos manuales.

## Qué crear en Supabase

Ejecuta estos archivos en el SQL Editor, en este orden:

1. [sql/001_base_schema.sql](/Users/stivenjohanhurtado/Integracion siigo nube/sql/001_base_schema.sql)
2. [sql/002_reporting_views.sql](/Users/stivenjohanhurtado/Integracion siigo nube/sql/002_reporting_views.sql)

## Variables necesarias

Completa estas variables en `.env` local y como secrets de GitHub:

- `SIIGO_USERNAME`
- `SIIGO_ACCESS_KEY`
- `SIIGO_PARTNER_ID`
- `SIIGO_BASE_URL`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

Opcionales:

- `SIIGO_AUTHORIZATION_HEADER`
- `SIIGO_CA_FILE`
- `SUPABASE_SCHEMA`
- `SIIGO_SYNC_LOOKBACK_DAYS`
- `SYNC_FROM_DATE`
- `SYNC_TO_DATE`

## Cómo correr localmente

Primera corrida acotada:

```bash
python3 sync_siigo_to_supabase.py --from-date 2026-04-01 --to-date 2026-04-30
```

Corrida incremental:

```bash
python3 sync_siigo_to_supabase.py
```

## Política de sincronización sugerida

Usa una ventana incremental amplia, por ejemplo `90` días.

Razón:

- una factura puede cambiar de saldo después de emitida
- si solo sincronizas 1 o 7 días por fecha de factura, podrías dejar saldos viejos desactualizados

Recomendación práctica:

- job diario con `SIIGO_SYNC_LOOKBACK_DAYS=90`
- backfill manual por rangos para cargar histórico inicial

## Qué consume Looker Studio

Conecta Looker Studio a estas vistas:

- `rpt_siigo_billing_by_customer`
- `rpt_siigo_billing_by_seller`
- `rpt_siigo_billing_by_customer_seller`
- `rpt_siigo_billing_by_day`
- `rpt_siigo_sync_health`

## Qué hace el workflow

El workflow [siigo_supabase_sync.yml](/Users/stivenjohanhurtado/Integracion siigo nube/.github/workflows/siigo_supabase_sync.yml):

- corre todos los días a las `11:15 UTC`
- permite ejecución manual con rango opcional
- usa los secrets de GitHub para no guardar credenciales en el repo
