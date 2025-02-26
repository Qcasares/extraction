# Word Document Analysis & Restructuring System

A Python tool that extracts, analyzes, and restructures content from Word documents.

## Features

- Recursively searches directories for Word documents (.docx)
- Extracts text content and metadata from documents
- Identifies document structure (headings, sections, etc.)
- Extracts named entities using natural language processing
- Optionally uses Ollama AI to enhance analysis
- Exports structured data in JSON, XML, or CSV formats

## Installation

### Regular installation

```bash
pip install -e .
```

### Development installation

```bash
pip install -e ".[dev]"
```

## Usage

### Basic usage

```bash
# Process all Word documents in a directory
python -m extraction --input-dir /path/to/documents --output-dir /path/to/output

# Specify output format (json, xml, or csv)
python -m extraction --input-dir /path/to/documents --format xml

# Enable verbose logging
python -m extraction --input-dir /path/to/documents --verbose
```

### Using with Ollama AI

First, ensure you have Ollama running locally or accessible via API. Then:

```bash
# Use default Ollama model (llama3)
python -m extraction --input-dir /path/to/documents --use-ai

# Specify a different Ollama model
python -m extraction --input-dir /path/to/documents --use-ai --model mistral

# Use a remote Ollama API
python -m extraction --input-dir /path/to/documents --use-ai --api-base "http://remote-server:11434/api"
```

### Help

```bash
python -m extraction --help
```

## Library Usage

You can also use the tool programmatically:

```python
from extraction.document_processor import DocumentProcessor

# Initialize the processor
processor = DocumentProcessor(
    input_directory="./documents",
    output_format="json",
    use_ai=True,
    ollama_model="llama3"
)

# Process documents
results = processor.process_all()

# Save structured output
processor.save_results("./output")
```

## Development

### Install development dependencies

```bash
pip install -e ".[dev]"
```

### Run tests

```bash
# Run tests
pytest

# Run tests with coverage report
pytest --cov=extraction
```

### Run formatters

```bash
black .
isort .
```

### Run type checker

```bash
mypy src tests
```

## Output Formats

### JSON

- Individual JSON files for each document
- Summary JSON file with all documents

### XML

- Individual XML files for each document with metadata, sections, and entities

### CSV

- Separate CSV files for metadata, sections, and entities
- Relationships maintained through document and section references