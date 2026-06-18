#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib import error, parse, request

from siigo_core import (
    SiigoApiError,
    build_siigo_config,
    customer_display_name,
    fetch_customer_details,
    fetch_paginated_results,
    fetch_users_map,
    json_request,
    load_dotenv,
    money_to_float,
    require_env,
    safe_text,
    to_decimal,
    user_display_name,
)


DEFAULT_LOOKBACK_DAYS = 90
DEFAULT_PAGE_SIZE = 100
DEFAULT_BATCH_SIZE = 250


class SupabaseSyncError(RuntimeError):
    pass


@dataclass
class SupabaseConfig:
    url: str
    service_role_key: str
    schema: str = "public"
    batch_size: int = DEFAULT_BATCH_SIZE


def build_supabase_config() -> SupabaseConfig:
    load_dotenv(Path(".env"))
    return SupabaseConfig(
        url=require_env("SUPABASE_URL").rstrip("/"),
        service_role_key=require_env("SUPABASE_SERVICE_ROLE_KEY"),
        schema=os.getenv("SUPABASE_SCHEMA", "public").strip() or "public",
        batch_size=int(os.getenv("SUPABASE_BATCH_SIZE", str(DEFAULT_BATCH_SIZE))),
    )


def parse_iso_date(value: str, label: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise SupabaseSyncError(f"{label} debe tener formato YYYY-MM-DD") from exc


def resolve_sync_range(args: argparse.Namespace) -> tuple[date, date]:
    env_from = os.getenv("SYNC_FROM_DATE", "").strip()
    env_to = os.getenv("SYNC_TO_DATE", "").strip()
    from_raw = args.from_date or env_from or None
    to_raw = args.to_date or env_to or None

    end_date = parse_iso_date(to_raw, "--to-date") if to_raw else date.today()
    if from_raw:
        start_date = parse_iso_date(from_raw, "--from-date")
    else:
        lookback_days = args.days_back
        if lookback_days is None:
            lookback_days = int(os.getenv("SIIGO_SYNC_LOOKBACK_DAYS", str(DEFAULT_LOOKBACK_DAYS)))
        if lookback_days < 1:
            raise SupabaseSyncError("--days-back debe ser mayor o igual a 1")
        start_date = end_date - timedelta(days=lookback_days)

    if start_date > end_date:
        raise SupabaseSyncError("La fecha inicial no puede ser mayor a la fecha final.")
    return start_date, end_date


def date_bounds_to_rfc3339(start_date: date, end_date: date) -> tuple[str, str]:
    return (
        f"{start_date.isoformat()}T00:00:00Z",
        f"{end_date.isoformat()}T23:59:59Z",
    )


def supabase_headers(config: SupabaseConfig, prefer: str | None = None) -> dict[str, str]:
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "apikey": config.service_role_key,
        "Authorization": f"Bearer {config.service_role_key}",
        "Accept-Profile": config.schema,
        "Content-Profile": config.schema,
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def supabase_request(
    config: SupabaseConfig,
    method: str,
    endpoint: str,
    payload: dict[str, Any] | list[dict[str, Any]] | None = None,
    prefer: str | None = None,
    timeout: int = 60,
) -> dict[str, Any] | list[Any] | str | None:
    url = f"{config.url}/rest/v1/{endpoint.lstrip('/')}"
    try:
        _, body = json_request(
            method,
            url,
            supabase_headers(config, prefer=prefer),
            payload=payload,
            timeout=timeout,
        )
        return body
    except SiigoApiError as exc:
        raise SupabaseSyncError(str(exc)) from exc


def chunked(rows: list[dict[str, Any]], size: int) -> Iterable[list[dict[str, Any]]]:
    for start in range(0, len(rows), size):
        yield rows[start : start + size]


def insert_sync_run(
    config: SupabaseConfig,
    sync_mode: str,
    start_date: date,
    end_date: date,
    metadata: dict[str, Any],
) -> int:
    body = supabase_request(
        config,
        "POST",
        "siigo_sync_runs?select=id",
        payload=[
            {
                "status": "running",
                "sync_mode": sync_mode,
                "from_date": start_date.isoformat(),
                "to_date": end_date.isoformat(),
                "metadata": metadata,
            }
        ],
        prefer="return=representation",
    )
    if not isinstance(body, list) or not body or not isinstance(body[0], dict) or "id" not in body[0]:
        raise SupabaseSyncError("No fue posible crear el registro de corrida en Supabase.")
    return int(body[0]["id"])


def update_sync_run(config: SupabaseConfig, run_id: int, payload: dict[str, Any]) -> None:
    supabase_request(
        config,
        "PATCH",
        f"siigo_sync_runs?id=eq.{run_id}",
        payload=payload,
        prefer="return=minimal",
    )


def upsert_rows(
    config: SupabaseConfig,
    table: str,
    rows: list[dict[str, Any]],
    conflict_columns: list[str],
) -> int:
    if not rows:
        return 0

    for batch in chunked(rows, config.batch_size):
        supabase_request(
            config,
            "POST",
            f"{table}?on_conflict={parse.quote(','.join(conflict_columns), safe=',')}",
            payload=batch,
            prefer="resolution=merge-duplicates,return=minimal",
        )
    return len(rows)


def first_contact_email(customer_body: dict[str, Any]) -> str:
    contacts = customer_body.get("contacts")
    if isinstance(contacts, list):
        for contact in contacts:
            if isinstance(contact, dict):
                email = safe_text(contact.get("email"))
                if email:
                    return email
    return ""


def first_phone(customer_body: dict[str, Any]) -> str:
    phones = customer_body.get("phones")
    if isinstance(phones, list):
        for phone in phones:
            if isinstance(phone, dict):
                indicative = safe_text(phone.get("indicative"))
                number = safe_text(phone.get("number"))
                joined = " ".join(part for part in (indicative, number) if part)
                if joined:
                    return joined
    return ""


def build_customer_rows(
    customers_map: dict[str, dict[str, Any]],
    synced_at: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for customer_id, body in customers_map.items():
        address = body.get("address") if isinstance(body.get("address"), dict) else {}
        city = address.get("city") if isinstance(address.get("city"), dict) else {}
        rows.append(
            {
                "customer_id": customer_id,
                "identification": safe_text(body.get("identification")),
                "branch_office": safe_text(body.get("branch_office")) or "0",
                "customer_name": customer_display_name(
                    body,
                    safe_text(body.get("identification")) or customer_id,
                ),
                "commercial_name": safe_text(body.get("commercial_name")),
                "person_type": safe_text(body.get("person_type")),
                "active": bool(body.get("active")),
                "vat_responsible": bool(body.get("vat_responsible")),
                "email": first_contact_email(body),
                "phone": first_phone(body),
                "address_line": safe_text(address.get("address")),
                "city_name": safe_text(city.get("city_name")),
                "state_name": safe_text(city.get("state_name")),
                "country_name": safe_text(city.get("country_name")),
                "source_payload": body,
                "updated_at": synced_at,
            }
        )
    return rows


def build_seller_rows(
    users_map: dict[str, dict[str, Any]],
    synced_at: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for seller_id, body in users_map.items():
        rows.append(
            {
                "seller_id": seller_id,
                "seller_name": user_display_name(body, seller_id),
                "username": safe_text(body.get("username")),
                "email": safe_text(body.get("email")),
                "identification": safe_text(body.get("identification")),
                "active": bool(body.get("active")),
                "source_payload": body,
                "updated_at": synced_at,
            }
        )
    return rows


def build_invoice_rows(
    raw_invoices: list[dict[str, Any]],
    customers_map: dict[str, dict[str, Any]],
    users_map: dict[str, dict[str, Any]],
    run_id: int,
    synced_at: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for invoice in raw_invoices:
        customer = invoice.get("customer") if isinstance(invoice.get("customer"), dict) else {}
        customer_id = safe_text(customer.get("id")) or None
        customer_body = customers_map.get(customer_id or "")
        seller_id = safe_text(invoice.get("seller")) or None
        seller_body = users_map.get(seller_id or "")

        total_value = to_decimal(invoice.get("total"))
        balance_value = to_decimal(invoice.get("balance"))
        paid_value = total_value - balance_value

        invoice_date = safe_text(invoice.get("date")) or None
        rows.append(
            {
                "invoice_id": safe_text(invoice.get("id")),
                "document_id": safe_text(invoice.get("document", {}).get("id"))
                if isinstance(invoice.get("document"), dict)
                else "",
                "invoice_name": safe_text(invoice.get("name")),
                "prefix": safe_text(invoice.get("prefix")),
                "number": safe_text(invoice.get("number")),
                "invoice_date": invoice_date,
                "invoice_month": f"{invoice_date[:7]}-01" if invoice_date else None,
                "customer_id": customer_id,
                "customer_identification": safe_text(customer.get("identification")),
                "customer_branch_office": safe_text(customer.get("branch_office")) or "0",
                "customer_name": customer_display_name(
                    customer_body,
                    safe_text(customer.get("identification")) or customer_id or "Sin cliente",
                ),
                "seller_id": seller_id,
                "seller_name": user_display_name(seller_body, seller_id or "Sin vendedor"),
                "seller_identification": safe_text(seller_body.get("identification"))
                if isinstance(seller_body, dict)
                else "",
                "seller_email": safe_text(seller_body.get("email")) if isinstance(seller_body, dict) else "",
                "total_amount": money_to_float(total_value),
                "paid_amount": money_to_float(paid_value),
                "balance_amount": money_to_float(balance_value),
                "annulled": bool(invoice.get("annulled")),
                "currency_code": safe_text(invoice.get("currency", {}).get("code"))
                if isinstance(invoice.get("currency"), dict)
                else "",
                "source_created_at": safe_text(invoice.get("metadata", {}).get("created"))
                if isinstance(invoice.get("metadata"), dict)
                else "",
                "sync_run_id": run_id,
                "raw_payload": invoice,
                "updated_at": synced_at,
            }
        )

    return rows


def build_voucher_rows(
    raw_vouchers: list[dict[str, Any]],
    run_id: int,
    synced_at: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for v in raw_vouchers:
        vid = safe_text(v.get("id"))
        if not vid:
            continue
        vdate = safe_text(v.get("date")) or None
        customer = v.get("customer") if isinstance(v.get("customer"), dict) else {}
        rows.append({
            "voucher_id":             vid,
            "voucher_name":           safe_text(v.get("name")),
            "voucher_number":         safe_text(v.get("number")),
            "voucher_date":           vdate,
            "voucher_month":          f"{vdate[:7]}-01" if vdate else None,
            "customer_id":            safe_text(customer.get("id")) or None,
            "customer_identification": safe_text(customer.get("identification")),
            "voucher_type":           safe_text(v.get("type")),
            "source_created_at":      safe_text(v.get("metadata", {}).get("created"))
                                      if isinstance(v.get("metadata"), dict) else "",
            "sync_run_id":            run_id,
            "raw_payload":            v,
            "updated_at":             synced_at,
        })
    return rows


def build_voucher_item_rows(
    raw_vouchers: list[dict[str, Any]],
    synced_at: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for v in raw_vouchers:
        vid   = safe_text(v.get("id"))
        vdate = safe_text(v.get("date")) or None
        if not vid or not vdate:
            continue
        for item in v.get("items", []):
            due = item.get("due")
            if not isinstance(due, dict):
                continue
            prefix      = safe_text(due.get("prefix"))
            consecutive = due.get("consecutive")
            value       = to_decimal(item.get("value"))
            if not prefix or consecutive is None or value <= 0:
                continue
            rows.append({
                "voucher_id":     vid,
                "voucher_date":   vdate,
                "voucher_month":  f"{vdate[:7]}-01",
                "invoice_name":   f"{prefix}-{consecutive}",
                "invoice_prefix": prefix,
                "invoice_number": int(consecutive),
                "value":          money_to_float(value),
                "updated_at":     synced_at,
            })
    return rows


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sincroniza datos de Siigo Nube hacia Supabase usando solo lectura."
    )
    parser.add_argument("--from-date", help="Fecha inicial YYYY-MM-DD.")
    parser.add_argument("--to-date", help="Fecha final YYYY-MM-DD.")
    parser.add_argument(
        "--days-back",
        type=int,
        help="Días hacia atrás si no defines rango exacto. Default: SIIGO_SYNC_LOOKBACK_DAYS o 90.",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        help=f"Cantidad solicitada por página a Siigo. Default: {DEFAULT_PAGE_SIZE}",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        help="Límite opcional para pruebas.",
    )
    parser.add_argument(
        "--sync-mode",
        default="incremental_window",
        help="Etiqueta lógica de la corrida. Default: incremental_window",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    run_id: int | None = None
    supabase_config: SupabaseConfig | None = None

    try:
        if args.page_size < 1:
            raise SupabaseSyncError("--page-size debe ser mayor o igual a 1")
        if args.max_pages is not None and args.max_pages < 1:
            raise SupabaseSyncError("--max-pages debe ser mayor o igual a 1")

        siigo_config = build_siigo_config()
        supabase_config = build_supabase_config()
        start_date, end_date = resolve_sync_range(args)
        start_rfc3339, end_rfc3339 = date_bounds_to_rfc3339(start_date, end_date)
        synced_at = datetime.now(timezone.utc).isoformat()

        run_id = insert_sync_run(
            supabase_config,
            sync_mode=args.sync_mode,
            start_date=start_date,
            end_date=end_date,
            metadata={
                "page_size_requested": args.page_size,
                "max_pages": args.max_pages,
                "date_start_rfc3339": start_rfc3339,
                "date_end_rfc3339": end_rfc3339,
            },
        )

        # ── Recibos de caja ───────────────────────────────────────────────────
        voucher_window_from = start_date - timedelta(days=120)
        voucher_window_to   = end_date   + timedelta(days=30)
        voucher_start_rfc, voucher_end_rfc = date_bounds_to_rfc3339(
            voucher_window_from, voucher_window_to
        )
        fetched_vouchers = fetch_paginated_results(
            siigo_config,
            "v1/vouchers",
            {
                "page":          "1",
                "page_size":     str(args.page_size),
                "created_start": voucher_start_rfc,
                "created_end":   voucher_end_rfc,
            },
            max_pages=args.max_pages,
        )
        raw_vouchers = [v for v in fetched_vouchers["results"] if isinstance(v, dict)]

        # ── Facturas ──────────────────────────────────────────────────────────
        fetched = fetch_paginated_results(
            siigo_config,
            "v1/invoices",
            {
                "page": "1",
                "page_size": str(args.page_size),
                "date_start": start_rfc3339,
                "date_end": end_rfc3339,
            },
            max_pages=args.max_pages,
        )
        raw_invoices = [invoice for invoice in fetched["results"] if isinstance(invoice, dict)]

        customer_hints: dict[str, dict[str, str]] = {}
        for invoice in raw_invoices:
            customer = invoice.get("customer")
            if not isinstance(customer, dict):
                continue
            customer_id = safe_text(customer.get("id"))
            if not customer_id:
                continue
            customer_hints[customer_id] = {
                "identification": safe_text(customer.get("identification")),
                "branch_office": safe_text(customer.get("branch_office")),
            }

        customers_map, customer_failures = fetch_customer_details(siigo_config, customer_hints)
        users_map, users_stats = fetch_users_map(siigo_config)

        customer_rows = build_customer_rows(customers_map, synced_at)
        seller_rows = build_seller_rows(users_map, synced_at)
        invoice_rows = build_invoice_rows(
            raw_invoices,
            customers_map,
            users_map,
            run_id=run_id,
            synced_at=synced_at,
        )

        voucher_rows      = build_voucher_rows(raw_vouchers, run_id=run_id, synced_at=synced_at)
        voucher_item_rows = build_voucher_item_rows(raw_vouchers, synced_at=synced_at)

        upserted_customers = upsert_rows(
            supabase_config,
            "siigo_customers",
            customer_rows,
            ["customer_id"],
        )
        upserted_sellers = upsert_rows(
            supabase_config,
            "siigo_sellers",
            seller_rows,
            ["seller_id"],
        )
        upserted_invoices = upsert_rows(
            supabase_config,
            "siigo_invoices",
            invoice_rows,
            ["invoice_id"],
        )
        upserted_vouchers = upsert_rows(
            supabase_config,
            "siigo_vouchers",
            voucher_rows,
            ["voucher_id"],
        )
        # Los items se reescriben completos por voucher: borrar los existentes
        # del lote y reinsertar para evitar duplicados por línea.
        if voucher_item_rows:
            voucher_ids = list({r["voucher_id"] for r in voucher_item_rows})
            for batch in chunked(voucher_ids, config.batch_size if False else 200):
                ids_csv = ",".join(f'"{vid}"' for vid in batch)
                supabase_request(
                    supabase_config,
                    "DELETE",
                    f"siigo_voucher_items?voucher_id=in.({ids_csv})",
                    prefer="return=minimal",
                )
            for batch in chunked(voucher_item_rows, supabase_config.batch_size):
                supabase_request(
                    supabase_config,
                    "POST",
                    "siigo_voucher_items",
                    payload=batch,
                    prefer="return=minimal",
                )

        update_sync_run(
            supabase_config,
            run_id,
            {
                "status": "completed",
                "finished_at": datetime.now(timezone.utc).isoformat(),
                "invoices_fetched": len(raw_invoices),
                "customers_upserted": upserted_customers,
                "sellers_upserted": upserted_sellers,
                "invoices_upserted": upserted_invoices,
                "metadata": {
                    "pages_fetched": fetched["pages_fetched"],
                    "reported_total_results": fetched["reported_total_results"],
                    "customer_lookup_failures": customer_failures,
                    "users_pages_fetched": users_stats["pages_fetched"],
                    "users_reported_total": users_stats["reported_total_results"],
                    "vouchers_upserted": upserted_vouchers,
                    "voucher_items_inserted": len(voucher_item_rows),
                },
            },
        )

        print(f"sync_run_id={run_id}")
        print(f"window={start_date.isoformat()}..{end_date.isoformat()}")
        print(f"invoices_fetched={len(raw_invoices)}")
        print(f"customers_upserted={upserted_customers}")
        print(f"sellers_upserted={upserted_sellers}")
        print(f"invoices_upserted={upserted_invoices}")
        print(f"vouchers_upserted={upserted_vouchers}")
        print(f"voucher_items_inserted={len(voucher_item_rows)}")
        return 0
    except (SiigoApiError, SupabaseSyncError) as exc:
        if supabase_config is not None and run_id is not None:
            try:
                update_sync_run(
                    supabase_config,
                    run_id,
                    {
                        "status": "failed",
                        "finished_at": datetime.now(timezone.utc).isoformat(),
                        "error_message": str(exc),
                    },
                )
            except Exception:
                pass
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
