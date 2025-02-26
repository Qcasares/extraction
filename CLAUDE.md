# Project Information

## Commands

### Development Setup
```bash
# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install development dependencies
pip install -e ".[dev]"
```

### Testing
```bash
# Run all tests
pytest

# Run verbose tests
pytest -v

# Run tests with coverage
pytest --cov=extraction

# Generate HTML coverage report
pytest --cov=extraction --cov-report=html
```

### Formatting and Linting
```bash
# Format code
black .
isort .

# Type checking
mypy src tests

# Install type stubs if needed
pip install types-requests
```

## Project Structure
- `src/extraction/`: Main package code
- `tests/`: Test files
- `pyproject.toml`: Project configuration

## Code Style
- Use type annotations for all functions and methods
- Follow PEP 8 style guide
- Line length: 88 characters (Black default)
- Docstrings: Google style