#!/usr/bin/env bash
# Corre psql contra Supabase con los DB_* del .env de Mesa de Entradas (sin exponer la contraseña).
# Uso: bash scripts/psql_supabase.sh [argumentos de psql]
set -euo pipefail

ENV_FILE="${ENV_FILE:-$(dirname "$0")/../.env}"
PSQL="${PSQL:-/c/Program Files/PostgreSQL/17/bin/psql}"

leer() { grep -E "^$1 *=" "$ENV_FILE" | head -1 | sed -E 's/^[^=]*= *//; s/\r$//; s/^"(.*)"$/\1/'; }

# El host directo (db.<ref>.supabase.co) es solo IPv6: se entra por el pooler (IPv4),
# donde el usuario lleva el ref del proyecto: postgres.<ref>.
REF="$(leer DB_HOST | sed -E 's/^db\.([^.]+)\..*/\1/')"
export PGHOST="${POOLER_HOST:-$(leer POOLER_HOST)}" PGPORT=5432
export PGDATABASE="$(leer DB_NAME)" PGUSER="$(leer DB_USER).$REF" PGPASSWORD="$(leer DB_PASSWORD)"
export PGCONNECT_TIMEOUT=15 PGSSLMODE=require

exec "$PSQL" -v ON_ERROR_STOP=1 "$@"
