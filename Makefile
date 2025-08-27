# ===== Makefile (safe to commit) =====

# Default service names
SERVICE ?= api
DC      ?= docker compose

# Show help by default
.DEFAULT_GOAL := help

## start services (build if needed)
up: ## Start stack in background
	$(DC) up -d

## stop services
down: ## Stop stack
	$(DC) down

## rebuild the API image and restart it
rebuild: ## Rebuild API (no cache) and restart
	$(DC) build --no-cache $(SERVICE)
	$(DC) up -d $(SERVICE)

## tail API logs (last 10m)
logs: ## Tail API logs
	$(DC) logs -f --since=10m $(SERVICE)

## run deploy (pull latest, rebuild, up)
deploy: ## Pull git + docker and restart API
	./deploy.sh

## database backup (uses backups/pg_dump.sh)
backup: ## Create DB backup (gzip) and rotate old ones
	./backups/pg_dump.sh

## quick health check (local)
health: ## curl /healthz
	./health.sh

## run alembic upgrade (if using Alembic)
migrate: ## Run DB migrations inside API container
	$(DC) exec $(SERVICE) alembic upgrade head

## prune unused docker stuff (careful)
prune: ## Clean dangling images/containers
	docker image prune -f
	docker system prune -f

## Print help
help:
	@awk 'BEGIN {FS = ":.*##"; printf "\nUsage: make \033[36m<TARGET>\033[0m\n\nTargets:\n"} /^[a-zA-Z0-9_\-]+:.*?##/ { printf "  \033[36m%-12s\033[0m %s\n", $$1, $$2 }' $(MAKEFILE_LIST)
