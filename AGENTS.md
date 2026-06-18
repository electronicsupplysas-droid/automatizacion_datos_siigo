# AGENTS.md — Integración Siigo Nube

## Reglas generales
- Respuestas y comentarios en español
- Usar urllib, evitar requests/httpx (cero dependencias externas)
- Python 3.10+, type hints obligatorios en todo
- No introducir librerías sin declararlas en requirements.txt

## Estructura del proyecto
- `siigo_core.py` — Cliente base de Siigo (única fuente de verdad)
- `siigo_readonly_test.py` — CLI de pruebas, importa desde siigo_core (NO duplicar lógica)
- `sync_siigo_to_supabase.py` — Sincronizador a Supabase
- `sql/` — Esquemas y vistas de Supabase
- `docs/` — Documentación
- `.github/workflows/` — CI/CD

## Estilo de código
- `dataclass` para configs, excepciones personalizadas para errores de dominio
- Decimal para montos, float solo al serializar
- No silenciar excepciones sin logging
- Safe navigation: `body.get("key")` en lugar de `body["key"]`
- Nombres de variables y funciones en inglés, docstrings en español

## Seguridad
- Toda ruta a Siigo debe empezar con `v1/`
- No hardcodear credenciales
- .env en .gitignore
- Sin POST/PUT/DELETE contra Siigo (solo GET + auth)

## Al finalizar cada tarea
- Si encuentro problemas o riesgos, listarlos con posibles soluciones
- No solo reportar el error, sugerir cómo resolverlo

## Comandos comunes
- `python3 sync_siigo_to_supabase.py --from-date YYYY-MM-DD --to-date YYYY-MM-DD`
- `python3 siigo_readonly_test.py auth-test`
- `python3 siigo_readonly_test.py get v1/customers --param page=1 --param page_size=5`
