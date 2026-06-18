#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import ssl
import sys
import time
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from urllib import error, parse, request


DEFAULT_BASE_URL = "https://api.siigo.com"

# HTTP status codes que merecen reintento (errores transitorios)
_RETRYABLE_STATUS = frozenset({429, 500, 502, 503, 504})
# Segundos de espera entre reintentos sucesivos (hasta 3 reintentos)
_RETRY_BACKOFF = (5, 15, 30)
AUTH_PATH_CANDIDATES = ("/auth", "/v1/auth")
SYSTEM_CA_CANDIDATES = (
    "/etc/ssl/cert.pem",
    "/private/etc/ssl/cert.pem",
)
MONEY_PLACES = Decimal("0.01")


class SiigoApiError(RuntimeError):
    pass


@dataclass
class SiigoConfig:
    username: str
    access_key: str
    partner_id: str
    base_url: str = DEFAULT_BASE_URL
    auth_header: str | None = None


def load_dotenv(env_path: Path) -> None:
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        os.environ.setdefault(key, value)


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise SiigoApiError(f"Falta la variable de entorno obligatoria: {name}")
    return value


def build_siigo_config() -> SiigoConfig:
    load_dotenv(Path(".env"))
    return SiigoConfig(
        username=require_env("SIIGO_USERNAME"),
        access_key=require_env("SIIGO_ACCESS_KEY"),
        partner_id=require_env("SIIGO_PARTNER_ID"),
        base_url=os.getenv("SIIGO_BASE_URL", DEFAULT_BASE_URL).rstrip("/"),
        auth_header=os.getenv("SIIGO_AUTHORIZATION_HEADER", "").strip() or None,
    )


def build_ssl_context() -> ssl.SSLContext:
    custom_cafile = os.getenv("SIIGO_CA_FILE", "").strip()
    if custom_cafile:
        cafile = Path(custom_cafile)
        if not cafile.exists():
            raise SiigoApiError(f"El archivo SIIGO_CA_FILE no existe: {cafile}")
        return ssl.create_default_context(cafile=str(cafile))

    default_paths = ssl.get_default_verify_paths()
    if default_paths.cafile and Path(default_paths.cafile).exists():
        return ssl.create_default_context()

    for candidate in SYSTEM_CA_CANDIDATES:
        if Path(candidate).exists():
            return ssl.create_default_context(cafile=candidate)

    return ssl.create_default_context()


def decode_body(raw_body: bytes, content_type: str) -> dict[str, Any] | list[Any] | str | None:
    if not raw_body:
        return None

    text = raw_body.decode("utf-8", errors="replace")
    if "application/json" in content_type.lower():
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            return text
    return text


def format_http_error(status: int, url: str, body: Any) -> str:
    if isinstance(body, (dict, list)):
        serialized = json.dumps(body, ensure_ascii=False)
    elif body is None:
        serialized = "(sin cuerpo)"
    else:
        serialized = str(body)
    return f"Error HTTP {status} al consumir {url}: {serialized}"


def json_request(
    method: str,
    url: str,
    headers: dict[str, str],
    payload: dict[str, Any] | list[Any] | None = None,
    timeout: int = 30,
) -> tuple[int, dict[str, Any] | list[Any] | str | None]:
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    ssl_context = build_ssl_context()

    for attempt in range(len(_RETRY_BACKOFF) + 1):
        req = request.Request(url=url, data=data, method=method.upper())
        for key, value in headers.items():
            req.add_header(key, value)

        try:
            with request.urlopen(req, timeout=timeout, context=ssl_context) as response:
                return response.status, decode_body(response.read(), response.headers.get("Content-Type", ""))

        except error.HTTPError as exc:
            if exc.code in _RETRYABLE_STATUS and attempt < len(_RETRY_BACKOFF):
                ra = (exc.headers.get("Retry-After") or "").strip()
                wait = int(ra) if ra.isdigit() else _RETRY_BACKOFF[attempt]
                print(
                    f"  ↻ HTTP {exc.code} — reintentando en {wait}s "
                    f"(intento {attempt + 1}/{len(_RETRY_BACKOFF)})…",
                    file=sys.stderr,
                )
                time.sleep(wait)
                continue
            raw_body = exc.read()
            body = decode_body(raw_body, exc.headers.get("Content-Type", ""))
            raise SiigoApiError(format_http_error(exc.code, url, body)) from exc

        except error.URLError as exc:
            if attempt < len(_RETRY_BACKOFF):
                wait = _RETRY_BACKOFF[attempt]
                print(
                    f"  ↻ Error de red — reintentando en {wait}s "
                    f"(intento {attempt + 1}/{len(_RETRY_BACKOFF)})…",
                    file=sys.stderr,
                )
                time.sleep(wait)
                continue
            raise SiigoApiError(f"No fue posible conectarse a Siigo: {exc.reason}") from exc

    raise SiigoApiError("Se agotaron los reintentos de conexión a Siigo.")


def obtain_token(config: SiigoConfig, timeout: int = 30) -> tuple[str, str]:
    payload = {
        "username": config.username,
        "access_key": config.access_key,
    }

    last_error: SiigoApiError | None = None
    for auth_path in AUTH_PATH_CANDIDATES:
        url = f"{config.base_url}{auth_path}"
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Partner-Id": config.partner_id,
        }
        if config.auth_header:
            headers["Authorization"] = config.auth_header

        try:
            _, body = json_request("POST", url, headers, payload=payload, timeout=timeout)
        except SiigoApiError as exc:
            last_error = exc
            continue

        if not isinstance(body, dict) or "access_token" not in body:
            raise SiigoApiError(
                f"La respuesta de autenticación no tiene el formato esperado: {body!r}"
            )

        token = str(body["access_token"]).strip()
        token_type = str(body.get("token_type", "Bearer")).strip() or "Bearer"
        if not token:
            raise SiigoApiError("Siigo respondió sin access_token.")
        return token, token_type

    if last_error is not None:
        raise last_error
    raise SiigoApiError("No fue posible obtener el token de autenticación.")


def normalize_path(path: str) -> str:
    cleaned = path.strip().lstrip("/")
    if "://" in cleaned:
        raise SiigoApiError("El path debe ser relativo, por ejemplo: v1/customers")
    if cleaned in {"auth", "v1/auth"}:
        raise SiigoApiError("La autenticación ya la maneja el cliente internamente.")
    if not cleaned.startswith("v1/"):
        raise SiigoApiError("Por seguridad el path debe iniciar con 'v1/'.")
    return cleaned


def authorization_candidates(token: str, token_type: str) -> list[str]:
    candidates: list[str] = []
    typed = token if " " in token else f"{token_type} {token}".strip()
    raw = token

    for candidate in (typed, raw):
        if candidate and candidate not in candidates:
            candidates.append(candidate)
    return candidates


def build_get_url(config: SiigoConfig, path: str, params: dict[str, str]) -> str:
    query = parse.urlencode(params)
    url = f"{config.base_url}/{normalize_path(path)}"
    if query:
        url = f"{url}?{query}"
    return url


def get_read_headers(config: SiigoConfig, authorization: str) -> dict[str, str]:
    return {
        "Accept": "application/json",
        "Authorization": authorization,
        "Partner-Id": config.partner_id,
    }


def readonly_get_url(
    config: SiigoConfig,
    url: str,
    timeout: int = 30,
    token: str | None = None,
    token_type: str | None = None,
) -> tuple[int, dict[str, Any] | list[Any] | str | None]:
    if token is None or token_type is None:
        token, token_type = obtain_token(config, timeout=timeout)

    last_error: SiigoApiError | None = None
    for authorization in authorization_candidates(token, token_type):
        headers = get_read_headers(config, authorization)
        try:
            return json_request("GET", url, headers, timeout=timeout)
        except SiigoApiError as exc:
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    raise SiigoApiError("La consulta GET no se pudo completar.")


def readonly_get(
    config: SiigoConfig,
    path: str,
    params: dict[str, str],
    timeout: int = 30,
) -> tuple[int, dict[str, Any] | list[Any] | str | None]:
    token, token_type = obtain_token(config, timeout=timeout)
    url = build_get_url(config, path, params)
    return readonly_get_url(
        config,
        url,
        timeout=timeout,
        token=token,
        token_type=token_type,
    )


def extract_next_href(body: dict[str, Any]) -> str | None:
    for links_key in ("_links", "__links"):
        links = body.get(links_key)
        if not isinstance(links, dict):
            continue
        next_link = links.get("next")
        if not isinstance(next_link, dict):
            continue
        href = next_link.get("href")
        if isinstance(href, str) and href.strip():
            return href.strip()
    return None


def build_next_url_from_pagination(current_url: str, body: dict[str, Any]) -> str | None:
    pagination = body.get("pagination")
    results = body.get("results")
    if not isinstance(pagination, dict) or not isinstance(results, list):
        return None

    page = pagination.get("page")
    page_size = pagination.get("page_size")
    total_results = pagination.get("total_results")
    if not all(isinstance(value, int) for value in (page, page_size, total_results)):
        return None
    if page < 1 or page_size < 1 or total_results <= page * page_size:
        return None

    parsed = parse.urlsplit(current_url)
    query = dict(parse.parse_qsl(parsed.query, keep_blank_values=True))
    query["page"] = str(page + 1)
    query.setdefault("page_size", str(page_size))
    new_query = parse.urlencode(query)
    return parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, new_query, parsed.fragment))


def fetch_paginated_results(
    config: SiigoConfig,
    path: str,
    params: dict[str, str],
    timeout: int = 30,
    max_pages: int | None = None,
) -> dict[str, Any]:
    normalized_path = normalize_path(path)
    current_url = build_get_url(config, normalized_path, params)
    token, token_type = obtain_token(config, timeout=timeout)

    seen_urls: set[str] = set()
    items: list[Any] = []
    pages_fetched = 0
    reported_total_results: int | None = None

    while current_url:
        if current_url in seen_urls:
            raise SiigoApiError(f"Se detectó un ciclo de paginación en {current_url}")
        seen_urls.add(current_url)

        status, body = readonly_get_url(
            config,
            current_url,
            timeout=timeout,
            token=token,
            token_type=token_type,
        )
        if status != 200:
            raise SiigoApiError(f"Respuesta inesperada HTTP {status} en {current_url}")
        if not isinstance(body, dict):
            raise SiigoApiError("La exportación esperaba una respuesta JSON tipo objeto.")

        results = body.get("results")
        if not isinstance(results, list):
            raise SiigoApiError(
                "La exportación paginada solo funciona con endpoints que respondan con 'results'."
            )

        items.extend(results)
        pages_fetched += 1

        pagination = body.get("pagination")
        if (
            reported_total_results is None
            and isinstance(pagination, dict)
            and isinstance(pagination.get("total_results"), int)
        ):
            reported_total_results = pagination["total_results"]

        if max_pages is not None and pages_fetched >= max_pages:
            break

        current_url = extract_next_href(body) or build_next_url_from_pagination(current_url, body)

    return {
        "path": normalized_path,
        "params": params,
        "pages_fetched": pages_fetched,
        "reported_total_results": reported_total_results,
        "results": items,
    }


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


_PLACEHOLDER_NAMES: frozenset[str] = frozenset({
    "noaplica", "n/a", "na", "sinnombre", "ninguno", "sinregistro",
})


def _is_placeholder(text: str) -> bool:
    return text.lower().replace(" ", "") in _PLACEHOLDER_NAMES


def customer_display_name(customer_body: dict[str, Any] | None, fallback_name: str) -> str:
    if isinstance(customer_body, dict):
        raw_name = customer_body.get("name")
        if isinstance(raw_name, list):
            joined = " ".join(
                part.strip() for part in raw_name if isinstance(part, str) and part.strip()
            )
            if joined and not _is_placeholder(joined):
                return joined
        if isinstance(raw_name, str) and raw_name.strip() and not _is_placeholder(raw_name.strip()):
            return raw_name.strip()

        commercial_name = safe_text(customer_body.get("commercial_name"))
        if commercial_name and not _is_placeholder(commercial_name):
            return commercial_name

    return fallback_name


def invoice_exchange_rate(invoice: dict[str, Any]) -> Decimal:
    """Returns the COP/foreign-currency rate for an invoice (1 if already COP)."""
    currency = invoice.get("currency")
    if not isinstance(currency, dict):
        return Decimal("1")
    code = safe_text(currency.get("code")) or ""
    if not code or code.upper() == "COP":
        return Decimal("1")
    rate = to_decimal(currency.get("exchange_rate", 1))
    return rate if rate > 0 else Decimal("1")


def dedupe_repeated_phrase(text: str) -> str:
    words = [word for word in text.split() if word]
    if len(words) % 2 == 0 and words:
        midpoint = len(words) // 2
        if [word.casefold() for word in words[:midpoint]] == [
            word.casefold() for word in words[midpoint:]
        ]:
            return " ".join(words[:midpoint])
    return text


def user_display_name(user_body: dict[str, Any] | None, fallback_name: str) -> str:
    if isinstance(user_body, dict):
        first_name = safe_text(user_body.get("first_name"))
        last_name = safe_text(user_body.get("last_name"))
        if first_name and last_name and first_name.casefold() == last_name.casefold():
            return first_name
        full_name = dedupe_repeated_phrase(
            " ".join(part for part in (first_name, last_name) if part).strip()
        )
        if full_name:
            return full_name

        for key in ("username", "email", "identification"):
            value = safe_text(user_body.get(key))
            if value:
                return value

    return fallback_name


def fetch_customer_details(
    config: SiigoConfig,
    customer_ids: dict[str, dict[str, str]],
    timeout: int = 30,
) -> tuple[dict[str, dict[str, Any]], list[dict[str, str]]]:
    if not customer_ids:
        return {}, []

    token, token_type = obtain_token(config, timeout=timeout)
    resolved: dict[str, dict[str, Any]] = {}
    failures: list[dict[str, str]] = []

    for customer_id, hints in sorted(customer_ids.items()):
        url = f"{config.base_url}/{normalize_path(f'v1/customers/{customer_id}')}"
        try:
            _, body = readonly_get_url(
                config,
                url,
                timeout=timeout,
                token=token,
                token_type=token_type,
            )
            if isinstance(body, dict):
                resolved[customer_id] = body
                continue
            raise SiigoApiError("Respuesta inesperada al consultar cliente.")
        except SiigoApiError as exc:
            failures.append(
                {
                    "customer_id": customer_id,
                    "customer_identification": hints.get("identification", ""),
                    "error": str(exc),
                }
            )

    return resolved, failures


def fetch_users_map(
    config: SiigoConfig,
    timeout: int = 30,
) -> tuple[dict[str, dict[str, Any]], dict[str, int]]:
    fetched = fetch_paginated_results(
        config,
        "v1/users",
        {
            "page": "1",
            "page_size": "100",
        },
        timeout=timeout,
    )

    users_map: dict[str, dict[str, Any]] = {}
    for user in fetched["results"]:
        if not isinstance(user, dict):
            continue
        user_id = safe_text(user.get("id"))
        if user_id:
            users_map[user_id] = user

    return users_map, {
        "pages_fetched": fetched["pages_fetched"],
        "reported_total_results": fetched["reported_total_results"] or len(users_map),
    }


def fetch_credit_notes(
    config: SiigoConfig,
    from_date: date,
    to_date: date,
    timeout: int = 30,
) -> list[dict[str, Any]]:
    """Descarga todas las notas crédito cuya fecha esté en [from_date, to_date]."""
    fetched = fetch_paginated_results(
        config,
        "v1/credit-notes",
        {
            "page": "1",
            "page_size": "100",
            "date_start": f"{from_date.isoformat()}T00:00:00Z",
            "date_end": f"{to_date.isoformat()}T23:59:59Z",
        },
        timeout=timeout,
    )
    from_str = from_date.isoformat()
    to_str   = to_date.isoformat()
    return [
        r for r in fetched["results"]
        if isinstance(r, dict)
        and from_str <= (safe_text(r.get("date")) or "")[:10] <= to_str
    ]


def to_decimal(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, str) and value.strip():
        try:
            return Decimal(value.strip())
        except InvalidOperation:
            return Decimal("0")
    return Decimal("0")


def money_to_float(value: Decimal) -> float:
    return float(value.quantize(MONEY_PLACES, rounding=ROUND_HALF_UP))
