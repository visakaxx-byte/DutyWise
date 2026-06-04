.PHONY: install run test validate-package clean

PYTHON ?= python3
PORT ?= 8002

install:
	$(PYTHON) -m venv .venv
	.venv/bin/pip install -r requirements.txt

run:
	.venv/bin/python -m uvicorn app:app --host 0.0.0.0 --port $(PORT)

test:
	PYTHONPYCACHEPREFIX=/tmp/bparty-v2-pycache $(PYTHON) -m unittest discover -s tests

validate-package:
	PYTHONPYCACHEPREFIX=/tmp/bparty-v2-pycache $(PYTHON) -m py_compile app.py config.py crawler_client.py engine.py llm_client.py run_example.py validate_live_flow.py

clean:
	rm -rf __pycache__ .pytest_cache outputs/validation
