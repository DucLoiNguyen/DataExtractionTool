.PHONY: install lint test run docker-build docker-run compose-up

install:
	pip install -r requirements-dev.txt

lint:
	ruff check

test:
	pytest -q

run:
	uvicorn webapp.app:app --reload

docker-build:
	docker build -t dataextraction-web:local .

docker-run:
	docker run --rm -p 8000:8000 dataextraction-web:local

compose-up:
	docker compose up --build
