create or replace view public.rpt_siigo_billing_by_customer as
select
    i.invoice_month,
    i.customer_id,
    coalesce(i.customer_name, c.customer_name) as customer_name,
    coalesce(i.customer_identification, c.identification) as customer_identification,
    coalesce(i.customer_branch_office, c.branch_office) as customer_branch_office,
    count(*)::integer as invoice_count,
    count(distinct nullif(i.seller_id, ''))::integer as seller_count,
    sum(i.total_amount)::numeric(18, 2) as total_sales,
    avg(i.total_amount)::numeric(18, 2) as avg_invoice_total,
    sum(i.paid_amount)::numeric(18, 2) as paid_amount,
    sum(i.balance_amount)::numeric(18, 2) as outstanding_balance
from public.siigo_invoices i
left join public.siigo_customers c on c.customer_id = i.customer_id
where not i.annulled
group by
    i.invoice_month,
    i.customer_id,
    coalesce(i.customer_name, c.customer_name),
    coalesce(i.customer_identification, c.identification),
    coalesce(i.customer_branch_office, c.branch_office);

create or replace view public.rpt_siigo_billing_by_seller as
select
    i.invoice_month,
    nullif(i.seller_id, '') as seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor') as seller_name,
    coalesce(i.seller_identification, s.identification) as seller_identification,
    coalesce(i.seller_email, s.email) as seller_email,
    count(*)::integer as invoice_count,
    count(distinct i.customer_id)::integer as customer_count,
    sum(i.total_amount)::numeric(18, 2) as total_sales,
    avg(i.total_amount)::numeric(18, 2) as avg_invoice_total,
    sum(i.paid_amount)::numeric(18, 2) as paid_amount,
    sum(i.balance_amount)::numeric(18, 2) as outstanding_balance
from public.siigo_invoices i
left join public.siigo_sellers s on s.seller_id = i.seller_id
where not i.annulled
group by
    i.invoice_month,
    nullif(i.seller_id, ''),
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor'),
    coalesce(i.seller_identification, s.identification),
    coalesce(i.seller_email, s.email);

create or replace view public.rpt_siigo_billing_by_customer_seller as
select
    i.invoice_month,
    i.customer_id,
    coalesce(i.customer_name, c.customer_name) as customer_name,
    coalesce(i.customer_identification, c.identification) as customer_identification,
    coalesce(i.customer_branch_office, c.branch_office) as customer_branch_office,
    nullif(i.seller_id, '') as seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor') as seller_name,
    coalesce(i.seller_identification, s.identification) as seller_identification,
    coalesce(i.seller_email, s.email) as seller_email,
    count(*)::integer as invoice_count,
    sum(i.total_amount)::numeric(18, 2) as total_sales,
    avg(i.total_amount)::numeric(18, 2) as avg_invoice_total,
    sum(i.paid_amount)::numeric(18, 2) as paid_amount,
    sum(i.balance_amount)::numeric(18, 2) as outstanding_balance
from public.siigo_invoices i
left join public.siigo_customers c on c.customer_id = i.customer_id
left join public.siigo_sellers s on s.seller_id = i.seller_id
where not i.annulled
group by
    i.invoice_month,
    i.customer_id,
    coalesce(i.customer_name, c.customer_name),
    coalesce(i.customer_identification, c.identification),
    coalesce(i.customer_branch_office, c.branch_office),
    nullif(i.seller_id, ''),
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor'),
    coalesce(i.seller_identification, s.identification),
    coalesce(i.seller_email, s.email);

create or replace view public.rpt_siigo_billing_by_day as
select
    i.invoice_date,
    i.invoice_month,
    count(*)::integer as invoice_count,
    count(distinct i.customer_id)::integer as customer_count,
    count(distinct nullif(i.seller_id, ''))::integer as seller_count,
    sum(i.total_amount)::numeric(18, 2) as total_sales,
    avg(i.total_amount)::numeric(18, 2) as avg_invoice_total,
    sum(i.paid_amount)::numeric(18, 2) as paid_amount,
    sum(i.balance_amount)::numeric(18, 2) as outstanding_balance
from public.siigo_invoices i
where not i.annulled
group by i.invoice_date, i.invoice_month;

create or replace view public.rpt_siigo_cartera as
select
    i.customer_id,
    coalesce(i.customer_name, c.customer_name) as customer_name,
    coalesce(i.customer_identification, c.identification) as customer_identification,
    coalesce(i.customer_branch_office, c.branch_office) as customer_branch_office,
    count(*)::integer as invoice_count,
    sum(i.balance_amount)::numeric(18, 2) as total_balance,
    sum(case when i.invoice_date >= current_date - 30 then i.balance_amount else 0 end)::numeric(18, 2) as current_30,
    sum(case when i.invoice_date between current_date - 60 and current_date - 31 then i.balance_amount else 0 end)::numeric(18, 2) as overdue_31_60,
    sum(case when i.invoice_date between current_date - 90 and current_date - 61 then i.balance_amount else 0 end)::numeric(18, 2) as overdue_61_90,
    sum(case when i.invoice_date < current_date - 90 then i.balance_amount else 0 end)::numeric(18, 2) as overdue_91_plus,
    max(i.invoice_date) as last_invoice_date
from public.siigo_invoices i
left join public.siigo_customers c on c.customer_id = i.customer_id
where not i.annulled and i.balance_amount > 0
group by
    i.customer_id,
    coalesce(i.customer_name, c.customer_name),
    coalesce(i.customer_identification, c.identification),
    coalesce(i.customer_branch_office, c.branch_office);

-- Comisiones por vendedor por mes de cobro real (fecha del recibo de caja)
create or replace view public.rpt_siigo_comisiones_por_mes as
select
    vi.voucher_month                                            as mes_cobro,
    i.seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name,
             'Sin vendedor')                                    as vendedor,
    count(distinct vi.voucher_id)::integer                      as recibos,
    count(*)::integer                                           as lineas,
    sum(vi.value)::numeric(18, 2)                               as monto_cobrado,
    (sum(vi.value) * 0.03)::numeric(18, 2)                      as comision_3pct
from public.siigo_voucher_items vi
join public.siigo_invoices i
    on i.invoice_name = vi.invoice_name
left join public.siigo_sellers s
    on s.seller_id = i.seller_id
where vi.value > 0
  and i.seller_id is not null
group by
    vi.voucher_month,
    i.seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor');


-- Comisiones acumuladas por vendedor (total del período en base de datos)
create or replace view public.rpt_siigo_comisiones_resumen as
select
    i.seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name,
             'Sin vendedor')                                    as vendedor,
    count(distinct vi.voucher_id)::integer                      as recibos,
    count(*)::integer                                           as lineas,
    sum(vi.value)::numeric(18, 2)                               as monto_cobrado,
    (sum(vi.value) * 0.03)::numeric(18, 2)                      as comision_3pct
from public.siigo_voucher_items vi
join public.siigo_invoices i
    on i.invoice_name = vi.invoice_name
left join public.siigo_sellers s
    on s.seller_id = i.seller_id
where vi.value > 0
  and i.seller_id is not null
group by
    i.seller_id,
    coalesce(nullif(i.seller_name, ''), s.seller_name, 'Sin vendedor')
order by comision_3pct desc;


create or replace view public.rpt_siigo_sync_health as
select
    r.id,
    r.started_at,
    r.finished_at,
    r.status,
    r.sync_mode,
    r.from_date,
    r.to_date,
    r.invoices_fetched,
    r.customers_upserted,
    r.sellers_upserted,
    r.invoices_upserted,
    r.error_message,
    r.metadata
from public.siigo_sync_runs r
order by r.started_at desc;
