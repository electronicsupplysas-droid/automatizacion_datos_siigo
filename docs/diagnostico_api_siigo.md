# Diagnóstico API Siigo Nube — Capacidades y Limitaciones

**Elaborado por:** Integración automatizada Electronic Supply S.A.S  
**Fecha:** Junio 2026  
**Versión API:** Siigo Nube v1  

---

## 1. Resumen ejecutivo

La API pública de Siigo Nube permite extraer información transaccional suficiente para construir informes de ventas, cartera y comisiones de forma automatizada. Sin embargo, presenta limitaciones importantes en cuanto a tipos de documento accesibles y comportamientos no documentados que requieren tratamiento especial para garantizar cifras correctas.

---

## 2. Endpoints disponibles y capacidades

### 2.1 Facturas de venta — `GET /v1/invoices`

| Campo | Disponible | Notas |
|---|---|---|
| Número de documento | ✅ | Campo `name` (ej. `FV-1-1234`) |
| Fecha de emisión | ✅ | Campo `date` |
| Fecha de vencimiento | ⚠️ | **No está en el raíz** — se extrae de `payments[].due_date` (máximo) |
| Cliente (NIT, nombre) | ✅ | Campo `customer.identification` / `customer.id` |
| Vendedor | ✅ | Campo `seller` (ID entero, requiere cruce con `/v1/users`) |
| Centro de costo | ✅ | Campo `cost_center` (ID entero, requiere cruce con `/v1/cost-centers`) |
| Moneda | ✅ | Campo `currency.code` (`COP` o `USD`) |
| Tasa de cambio | ✅ | Campo `currency.exchange_rate` (fija al momento de emisión) |
| Total | ✅ | Campo `total` — usar este valor directamente (ver §4.1) |
| Saldo pendiente | ✅ | Campo `balance` — saldo real sin cobrar |
| Items / líneas de detalle | ✅ | Array `items[]` con producto, cantidad, precio, impuestos |
| Estado anulado | ✅ | Campo `annulled: true/false` |
| Paginación | ✅ | Parámetros `page` y `page_size` (máx. 100 por página) |
| Filtro por fecha | ✅ | `date_start` / `date_end` en formato RFC3339 |

**Tipos de FV accesibles:**
- `FV` estándar — ID tipo documento 27511 / 27515
- Facturas en USD para cliente D1 S.A.S — accesibles con conversión automática

---

### 2.2 Notas crédito — `GET /v1/credit-notes`

| Campo | Disponible | Notas |
|---|---|---|
| Número de documento | ✅ | Campo `name` (ej. `NC-2-266`) |
| Factura referenciada | ✅ | Campo `invoice.name` |
| Total | ✅ | Campo `total` |
| Cliente, vendedor, fecha | ✅ | Misma estructura que facturas |
| Saldo pendiente | ❌ | Las NC no tienen campo `balance` |

---

### 2.3 Recibos de caja — `GET /v1/vouchers`

| Campo | Disponible | Notas |
|---|---|---|
| Número de recibo | ✅ | Campo `name` (ej. `RC-2-123`) |
| Fecha de cobro | ✅ | Campo `date` |
| Factura pagada | ✅ | `items[].due.prefix` + `items[].due.consecutive` |
| Monto aplicado | ✅ | Campo `items[].value` |
| Cliente, vendedor | ✅ | Misma estructura que facturas |
| Filtro por fecha | ✅ | `created_start` / `created_end` |

> **Nota importante:** Un recibo puede pagar varias facturas. Siempre iterar `items[]` y no asumir un RC = una FV.

---

### 2.4 Clientes — `GET /v1/customers`

| Campo | Disponible | Notas |
|---|---|---|
| NIT / identificación | ✅ | Campo `identification` |
| Nombre comercial | ✅ | Campo `name` |
| Ciudad | ✅ | Campo `address.city.city_name` (anidado 3 niveles) |
| Correo, teléfono | ✅ | Campos `contacts[]` |
| Paginación | ✅ | Igual que facturas |

---

### 2.5 Vendedores — `GET /v1/users`

| Campo | Disponible | Notas |
|---|---|---|
| ID numérico | ✅ | Se cruza con campo `seller` de las facturas |
| Nombre completo | ✅ | Campos `first_name` + `last_name` |
| Email | ✅ | Campo `email` |
| Paginación | ✅ | |

---

### 2.6 Centros de costo — `GET /v1/cost-centers`

| Campo | Disponible | Notas |
|---|---|---|
| ID numérico | ✅ | Se cruza con campo `cost_center` de las facturas |
| Nombre | ✅ | |
| Paginación | ❌ | **Devuelve lista plana, no objeto paginado** — tratar diferente al resto |

---

### 2.7 Tipos de documento — `GET /v1/document-types`

Disponible para consulta. Útil para identificar qué tipos existen en la cuenta. No se usa en el flujo principal de extracción.

---

## 3. Flujo de autenticación

- **Endpoint:** `POST /auth` o `POST /v1/auth`
- **Credenciales:** `username` + `access_key` + `Partner-Id` (header)
- **Token:** Bearer JWT con expiración (se renueva automáticamente en cada ejecución)
- **Restricción de seguridad:** Solo se permite el POST de autenticación. Ningún endpoint de negocio admite POST/PUT/PATCH/DELETE en esta integración — es **solo lectura**.

---

## 4. Limitaciones encontradas

### 4.1 Phantom IVA — cálculo incorrecto desde items

**Problema:** Algunas líneas de factura tienen `price = 0` pero `taxes[].total > 0`. Calcular el total como `Σ(qty × price) + Σ(taxes)` produce un resultado incorrecto.

**Solución aplicada:** Usar siempre `doc.total` (campo oficial de Siigo) como valor del total del documento. El desglose por ítem es informativo pero no se usa para el total.

---

### 4.2 Fecha de vencimiento vacía en el raíz

**Problema:** El campo `due_date` en el raíz de una factura siempre está vacío (`null`), aunque la factura tenga fecha de vencimiento configurada.

**Solución aplicada:** Extraer la fecha de vencimiento desde `payments[].due_date` y tomar el máximo cuando hay múltiples pagos.

---

### 4.3 Documentos DF (Débito Facturación) — no accesibles

**Problema:** Los documentos de tipo DF (IDs 10258 / 27513 / 27517) existen en Siigo pero el endpoint `GET /v1/invoices` nunca los devuelve, independientemente de los filtros aplicados.

**Impacto:** El cliente **QUICENO & CIA S.C.A.** tiene documentos DF que no aparecen en la integración. Diferencia conocida y permanente: **$1,461,715 COP**.

**Estado:** Sin solución posible con la API pública actual.

---

### 4.4 Documentos ND (Nota Débito) — no accesibles

**Problema:** Los documentos de tipo ND (IDs 10259 / 30465) tampoco son accesibles vía `GET /v1/credit-notes` ni ningún otro endpoint.

**Impacto:** 30 documentos ND identificados en la cuenta que no aparecen en ningún informe.

**Estado:** Sin solución posible con la API pública actual.

---

### 4.5 Facturas USD sin vendedor asignado (FV-2)

**Problema:** Las facturas emitidas en USD para el cliente D1 S.A.S (NIT 900276962), tipo documento 27515, no tienen vendedor asignado en Siigo (`seller = null`).

**Impacto:** Aparecen como "Sin vendedor" en el informe de comisiones. El volumen es significativo (~$4,700M COP en 2026).

**Estado:** Decisión de negocio pendiente — asignar vendedor en Siigo o tratarlas como comisión no aplicable.

---

### 4.6 Límite de tasa de consultas (Rate limiting)

**Problema:** La API retorna `HTTP 429` cuando se hacen demasiadas solicitudes en poco tiempo.

**Solución aplicada:** Reintentos automáticos con backoff exponencial: 5s → 15s → 30s (hasta 3 reintentos). Respeta el header `Retry-After` cuando está presente.

---

### 4.7 Límite de rango de fechas por consulta

**Problema:** Consultas con rangos de más de un mes no devuelven resultados completos — la API omite meses intermedios.

**Solución aplicada:** El orquestador divide automáticamente cualquier rango multi-mes en consultas mensuales individuales.

---

### 4.8 Errores de red transitorios

**Problema:** Conexiones cortadas o timeouts intermitentes, especialmente en descargas largas de cartera (223+ documentos).

**Solución aplicada:** Los mismos reintentos automáticos del punto 4.6 aplican también para errores de red (`URLError`).

---

## 5. Tabla resumen de accesibilidad por tipo de documento

| Tipo | Descripción | Accesible vía API | Endpoint |
|---|---|---|---|
| FV | Factura de venta | ✅ | `/v1/invoices` |
| NC | Nota crédito | ✅ | `/v1/credit-notes` |
| RC | Recibo de caja | ✅ | `/v1/vouchers` |
| DF | Débito facturación | ❌ | No disponible |
| ND | Nota débito | ❌ | No disponible |

---

## 6. Precisión de los datos obtenidos

Validación realizada contra el reporte oficial de Siigo **"Ventas por cliente — Enero a Junio 2026"**:

| Métrica | Valor |
|---|---|
| Total Siigo (reporte oficial) | $9,908,215,773.28 |
| Total integración (Supabase) | $9,917,846,685.00 |
| Diferencia total | $9,630,912 (0.097%) |
| Diferencia por DF (QUICENO) | $1,461,715 — permanente, API no lo expone |
| Diferencia por corte horario (OXXO) | $8,169,197 — facturas creadas después del corte del reporte |
| **Diferencia estructural real** | **$1,461,715 (0.015%)** |

La integración reproduce el 99.98% del valor total de ventas. La única diferencia estructural es el documento DF de QUICENO, que la API pública de Siigo no expone bajo ninguna circunstancia.

---

## 7. Arquitectura de la solución implementada

```
GitHub Actions (cron L-V 6AM COL)
        │
        ▼
  run_informes.py
        │
        ├── v1/invoices      ─┐
        ├── v1/credit-notes  ─┤──► tabla: documentos
        ├── v1/customers     ─┤    (una fila por FV o NC)
        ├── v1/users         ─┘
        │
        ├── v1/vouchers ──────────► tabla: recibos_caja
        │                           (una fila por pago RC)
        │
        └── v1/cost-centers (lookup, no paginado)
                │
                ▼
           Supabase (PostgreSQL)
                │
                ▼
        Vistas para Looker Studio:
          v_ventas_mes
          v_cartera
          v_comisiones_cobradas
          v_comisiones_pendientes
```

---

*Documento generado a partir de la experiencia de integración directa con la API de Siigo Nube v1.*
