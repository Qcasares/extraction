# Word Document Analysis & Restructuring System

A Python tool that extracts, analyzes, and restructures content from Word documents.

## Features

- Recursively searches directories for Word documents (.docx)
- Extracts text content and metadata from documents
- Identifies document structure (headings, sections, etc.)
- Extracts named entities using natural language processing
- Optionally uses AI models (Ollama, OpenAI, Anthropic) to enhance analysis
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

### Using with AI Integration

The tool supports multiple AI providers for enhanced document analysis:

#### Ollama (Local)

First, ensure you have Ollama running locally or accessible via API. Then:

```bash
# Use default Ollama model (llama3)
python -m extraction --input-dir /path/to/documents --use-ai

# Specify a different Ollama model
python -m extraction --input-dir /path/to/documents --use-ai --model mistral

# Use a remote Ollama API
python -m extraction --input-dir /path/to/documents --use-ai --api-base "http://remote-server:11434/api"
```

#### OpenAI

```bash
# Use OpenAI (requires API key)
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --api-key "your-api-key"

# Specify a different model
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "gpt-4-turbo" --api-key "your-api-key"

# Use other OpenAI models
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "gpt-4o-mini" --api-key "your-api-key"
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "o1-preview" --api-key "your-api-key"
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "o1-mini" --api-key "your-api-key"
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "o3-preview" --api-key "your-api-key"
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider openai --model "o3-mini" --api-key "your-api-key"
```

#### Anthropic Claude

```bash
# Use Anthropic Claude (requires API key)
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider anthropic --api-key "your-api-key"

# Specify a different model
python -m extraction --input-dir /path/to/documents --use-ai --ai-provider anthropic --model "claude-3-sonnet-20240229" --api-key "your-api-key"
```

#### Advanced AI Options

```bash
# Customize AI features
python -m extraction --input-dir /path/to/documents --use-ai --ai-features "summary,topics,sentiment,insights"

# Adjust response parameters (all models use max-tokens parameter with this tool)
python -m extraction --input-dir /path/to/documents --use-ai --temperature 0.3 --max-tokens 4000

# Disable response caching
python -m extraction --input-dir /path/to/documents --use-ai --disable-ai-cache
```

### Help

```bash
python -m extraction --help
```

## Library Usage

You can also use the tool programmatically:

```python
from extraction.document_processor import DocumentProcessor

# Initialize the processor with Ollama
processor = DocumentProcessor(
    input_directory="./documents",
    output_format="json",
    use_ai=True,
    ai_provider="ollama",
    ai_model="llama3"
)

# Or initialize with OpenAI
processor = DocumentProcessor(
    input_directory="./documents",
    output_format="json",
    use_ai=True,
    ai_provider="openai",
    ai_model="gpt-4o",
    api_key="your-api-key"
)

# Configure specific AI features
processor = DocumentProcessor(
    input_directory="./documents",
    output_format="json",
    use_ai=True,
    ai_provider="anthropic",
    ai_model="claude-3-opus-20240229",
    api_key="your-api-key",
    ai_features="summary,topics,sentiment",
    temperature=0.2,
    max_tokens=3000
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