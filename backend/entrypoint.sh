#!/usr/bin/env bash
set -euo pipefail

# Load envs if present
if [ -f "/app/.env" ]; then set -a; source /app/.env; set +a; fi
if [ -f "/app/.env.prod" ]; then set -a; source /app/.env.prod; set +a; fi

# Wait for Postgres (compose will inject DB host as "db")
if [[ -n "${DATABASE_HOST:-}" ]]; then
  echo "Waiting for Postgres at ${DATABASE_HOST}:${DATABASE_PORT:-5432}..."
  until pg_isready -h "${DATABASE_HOST}" -p "${DATABASE_PORT:-5432}" -U "${DATABASE_USER:-postgres}" >/dev/null 2>&1; do
    sleep 1
  done
fi

# Run Alembic migrations if configured
if [ -f "/app/alembic.ini" ]; then
  echo "Running Alembic migrations..."
  alembic upgrade head
fi

# Start Gunicorn
APP_MODULE="${GUNICORN_APP:-wsgi:app}"   # override if your app module differs
WORKERS="${GUNICORN_WORKERS:-2}"
BIND="${GUNICORN_BIND:-0.0.0.0:8000}"
TIMEOUT="${GUNICORN_TIMEOUT:-30}"
KEEPALIVE="${GUNICORN_KEEPALIVE:-2}"

exec gunicorn "$APP_MODULE" \
  --workers "$WORKERS" \
  --bind "$BIND" \
  --timeout "$TIMEOUT" \
  --keep-alive "$KEEPALIVE" \
  --access-logfile '-' --error-logfile '-'
