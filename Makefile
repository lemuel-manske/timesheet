VENV ?= .venv

PYTHON ?= $(if $(wildcard $(VENV)/bin/python),$(VENV)/bin/python,python3)


.PHONY: install test fmt check build-web serve


install:
	uv sync --all-groups


test:
	uv run pytest


fmt:
	uv run ruff format .
	uv run ruff check --fix .

check:
	uv run ruff check .
	uv run pyright


build-web:
	mkdir -p web/python
	cp src/*.py web/python/
	cp labels.json web/python/

serve: build-web
	$(PYTHON) -m http.server 8000 --directory web
