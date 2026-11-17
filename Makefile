VENV ?= .venv

PYTHON ?= $(if $(wildcard $(VENV)/bin/python),$(VENV)/bin/python,python3)

init:
	python3 -m venv $(VENV)
	$(VENV)/bin/python -m pip install --upgrade pip
	$(MAKE) install PYTHON=$(VENV)/bin/python

install:
	$(PYTHON) -m pip install -r requirements.txt

test:
	$(PYTHON) -m pytest -v

build-web:
	mkdir -p web/python
	cp extraction.py core.py excel_adapter.py labels.json web/python/

serve: build-web
	$(PYTHON) -m http.server 8000 --directory web
