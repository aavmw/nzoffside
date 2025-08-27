test:
	cd backend && pytest --cov=app --cov-report=term-missing

lint:
	cd backend && ruff check app tests && black --check app tests

fmt:
	cd backend && black app tests

typecheck:
	cd backend && mypy app
