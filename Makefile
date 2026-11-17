VENV ?= .venv

PYTHON ?= $(if $(wildcard $(VENV)/bin/python),$(VENV)/bin/python,python3)

init:
	python3 -m venv $(VENV)
	$(VENV)/bin/python -m pip install --upgrade pip
	$(MAKE) install PYTHON=$(VENV)/bin/python

test:
	$(PYTHON) -m pytest -v
