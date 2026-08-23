.PHONY: help install dev test test-cov lint format typecheck build clean run-workspace

help:
	@echo "Docx-Agent V2.1 Developer Commands:"
	@echo "  make install        Install package dependencies"
	@echo "  make dev            Install package in editable mode with dev dependencies"
	@echo "  make test           Run pytest suite"
	@echo "  make test-cov       Run pytest with code coverage report"
	@echo "  make lint           Check code with Ruff"
	@echo "  make format         Auto-format code with Ruff"
	@echo "  make typecheck      Run mypy static type checking"
	@echo "  make build          Build wheel and sdist distribution packages"
	@echo "  make clean          Clean build artifacts, caches, and temporary files"
	@echo "  make run-workspace  Launch local visual workspace server"

install:
	pip install .

dev:
	pip install -e ".[dev,mcp]"

test:
	pytest tests -v

test-cov:
	pytest tests -v --cov=docx_agent --cov-report=term-missing --cov-report=html

lint:
	ruff check src tests

format:
	ruff format src tests
	ruff check --fix src tests

typecheck:
	mypy src --ignore-missing-imports

build: clean
	python -m build

clean:
	rm -rf build/ dist/ *.egg-info .pytest_cache htmlcov .coverage
	find . -type d -name "__pycache__" -exec rm -rf {} +

run-workspace:
	docx-agent workspace
