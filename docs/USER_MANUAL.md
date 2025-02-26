# Extraction - User Manual

## Table of Contents
1. [Introduction](#introduction)
2. [Installation](#installation)
3. [Basic Usage](#basic-usage)
4. [Command Line Interface](#command-line-interface)
5. [Advanced AI Analysis](#advanced-ai-analysis)
6. [Output Formats](#output-formats)
7. [Python API](#python-api)
8. [Troubleshooting](#troubleshooting)
9. [Best Practices](#best-practices)

## Introduction

Extraction is a powerful Python library for Word document analysis and restructuring. It scans directories for .docx files, extracts content and metadata, analyzes document structure, and exports results to structured formats (JSON, XML, CSV).

Key features include:
- Automatic extraction of document metadata (title, author, dates)
- Section identification based on heading styles
- Named entity recognition using spaCy
- Advanced AI analysis using multiple AI providers (Ollama, OpenAI, Anthropic)
- Flexible export options (JSON, XML, CSV)

## Installation

### Prerequisites
- Python 3.8 or later
- spaCy and its English language model
- For AI analysis (optional):
  - Ollama (for local AI models)
  - OpenAI API key (for using OpenAI models)
  - Anthropic API key (for using Claude models)

### Install from PyPI
```bash
pip install extraction
```

### Install from Source
```bash
# Clone the repository
git clone https://github.com/yourusername/extraction.git
cd extraction

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install package with development dependencies
pip install -e ".[dev]"

# Download spaCy model
python -m spacy download en_core_web_sm
```

### For AI Analysis

#### Ollama (Local)
To use local AI models, install Ollama following the instructions at [ollama.ai](https://ollama.ai). After installation, run:

```bash
ollama pull llama3  # or any other model you prefer
```

#### OpenAI
To use OpenAI models, you need an API key from [OpenAI](https://platform.openai.com/). Once you have your API key, you can use it with the `--api-key` option.

#### Anthropic
To use Anthropic Claude models, you need an API key from [Anthropic](https://www.anthropic.com/). Once you have your API key, you can use it with the `--api-key` option.

## Basic Usage

### Command Line
```bash
# Basic usage
extraction --input-dir /path/to/docx/files --output-dir /path/to/output

# Choose output format
extraction --input-dir /path/to/docx/files --format json
extraction --input-dir /path/to/docx/files --format xml
extraction --input-dir /path/to/docx/files --format csv

# Enable AI analysis with Ollama
extraction --input-dir /path/to/docx/files --use-ai

# Use OpenAI
extraction --input-dir /path/to/docx/files --use-ai --ai-provider openai --api-key "your-api-key"

# Use Anthropic
extraction --input-dir /path/to/docx/files --use-ai --ai-provider anthropic --api-key "your-api-key"
```

### Python
```python
from extraction.document_processor import DocumentProcessor

# Initialize processor with Ollama
processor = DocumentProcessor(
    input_directory="/path/to/docx/files",
    output_format="json",
    use_ai=True,
    ai_provider="ollama",
    ai_model="llama3"
)

# Or with OpenAI
processor = DocumentProcessor(
    input_directory="/path/to/docx/files",
    output_format="json",
    use_ai=True,
    ai_provider="openai",
    ai_model="gpt-4o",
    api_key="your-api-key"
)

# Process all documents
processed_docs = processor.process_all()

# Save results
processor.save_results("/path/to/output")
```

## Command Line Interface

### Basic Options
```
--input-dir, -i        Directory containing Word documents to process (required)
--output-dir, -o       Directory to save processed results (default: ./output)
--format, -f           Output format: json, xml, or csv (default: json)
--verbose, -v          Enable verbose logging
```

### AI Analysis Options
```
--use-ai               Enable AI analysis
--ai-provider          AI provider to use: ollama, openai, anthropic (default: ollama)
--model, -m            AI model to use (provider-specific defaults)
--api-base             Base URL for AI provider API (provider-specific defaults)
--api-key              API key for authentication (required for OpenAI and Anthropic)
--temperature          Temperature for AI generation (0.0-1.0, default: 0.1)
--max-tokens           Maximum tokens for AI response (default: 2000)
--ai-timeout           Timeout in seconds for AI API calls (default: 60)
--ai-features          Comma-separated list of AI features to enable (default: all)
--disable-ai-cache     Disable caching of AI responses
--ai-cache-size        Maximum number of AI responses to cache (default: 100)
```

#### Available AI Features
- `summary` - Document summarization
- `topics` - Key topic identification
- `categories` - Content categorization
- `sentiment` - Sentiment analysis
- `relationships` - Entity relationship mapping
- `quality` - Document quality scoring
- `suggestions` - Improvement suggestions
- `themes` - Theme analysis
- `insights` - Key insights extraction

Example:
```bash
# Enable only specific AI features with Ollama
extraction --input-dir /path/to/docx/files --use-ai --ai-features "summary,sentiment,suggestions"

# With OpenAI
extraction --input-dir /path/to/docx/files --use-ai --ai-provider openai --api-key "your-api-key" --ai-features "summary,sentiment,suggestions"
```

## Advanced AI Analysis

Extraction can perform advanced analysis of your documents using various AI models. This functionality provides deeper insights into document content, structure, and quality.

### Setup Requirements

#### For Ollama (Local/Self-Hosted)
1. Install Ollama from [ollama.ai](https://ollama.ai)
2. Download your preferred model:
   ```bash
   ollama pull llama3  # Recommended model
   ```
3. Start the Ollama service:
   ```bash
   ollama serve
   ```

#### For OpenAI
1. Create an account on [OpenAI](https://platform.openai.com/)
2. Generate an API key in your account settings
3. Use the API key with the `--api-key` option

#### For Anthropic Claude
1. Create an account on [Anthropic](https://www.anthropic.com/)
2. Generate an API key in your account settings
3. Use the API key with the `--api-key` option

### AI Analysis Features

#### Document Summarization
Generates a concise summary of the document content, focusing on key points and main ideas.

#### Topic Identification
Identifies the primary topics discussed in the document, helping to quickly understand document focus.

#### Content Categorization
Classifies document content into categories like "technical", "business", "academic", etc.

#### Sentiment Analysis
Analyzes the overall sentiment of the document (positive, negative, neutral, mixed).

#### Entity Relationship Mapping
Identifies relationships between entities mentioned in the document, useful for understanding connections between people, organizations, etc.

#### Document Quality Assessment
Provides a numerical score (0-10) rating the overall quality of the document, considering factors like clarity, structure, and completeness.

#### Improvement Suggestions
Offers actionable suggestions for improving the document's clarity, structure, or content.

#### Theme Analysis
Identifies major themes in the document and their relevance, including which sections contain each theme.

#### Key Insights Extraction
Extracts important insights and takeaways from the document content.

### AI Performance Tips

1. **Provider Selection**:
   - `ollama`: Good for local processing, privacy, and no API costs
   - `openai`: High quality results with larger token limits
   - `anthropic`: Excellent for complex document analysis and detailed insights

2. **Model Selection**:
   - Ollama models:
     - `llama3`: Well-balanced for most analysis tasks
     - `mistral`: May provide better quality for technical documents
     - `phi`: Works well for shorter, less complex documents
   - OpenAI models:
     - `gpt-4o`: Best quality for comprehensive analysis
     - `gpt-3.5-turbo`: Faster and more cost-effective
   - Anthropic models:
     - `claude-3-opus-20240229`: Best quality for comprehensive analysis
     - `claude-3-sonnet-20240229`: Good balance of quality and speed

3. **Temperature Setting**: Lower temperature (0.1-0.3) provides more consistent, focused analysis. Higher temperature (0.7-0.9) generates more creative but less predictable analysis.

4. **API Timeouts**: For large documents, increase the timeout value:
   ```bash
   extraction --input-dir /path/to/docx/files --use-ai --ai-timeout 120
   ```

5. **Response Caching**: Caching is enabled by default to improve performance when processing similar documents. Disable it for more dynamic results:
   ```bash
   extraction --input-dir /path/to/docx/files --use-ai --disable-ai-cache
   ```

## Output Formats

### JSON
The JSON format provides the most comprehensive output, including all extracted data in a structured format. For documents with AI analysis, this includes the complete analysis data.

Example structure:
```json
{
  "metadata": {
    "title": "Document Title",
    "author": "Document Author",
    "created_date": "2023-03-15T10:30:00",
    "last_modified_date": "2023-03-16T14:45:00",
    "file_path": "/absolute/path/to/document.docx",
    "filename": "document.docx"
  },
  "sections": [
    {
      "title": "Introduction",
      "level": 1,
      "content": "Section content...",
      "entities": [
        {
          "text": "Entity name",
          "label": "PERSON",
          "start": 45,
          "end": 56
        }
      ]
    }
  ],
  "ai_analysis": {
    "summary": "A concise summary of the document.",
    "key_topics": ["Topic 1", "Topic 2"],
    "sentiment": "positive",
    "document_quality_score": 8.5,
    "improvement_suggestions": ["Suggestion 1", "Suggestion 2"],
    // Other AI analysis data...
  }
}
```

### XML
The XML format provides a structured representation suitable for integration with XML-based systems.

### CSV
The CSV format generates multiple files for different types of data:
- `document_metadata_*.csv` - Document metadata
- `document_sections_*.csv` - Section data
- `document_entities_*.csv` - Named entity data
- `document_ai_analysis_*.csv` - AI analysis summary
- `document_ai_topics_*.csv` - Topics, categories, insights, and suggestions
- `document_ai_relationships_*.csv` - Entity relationships
- `document_ai_themes_*.csv` - Theme analysis

## Python API

### DocumentProcessor

The main class for processing documents.

```python
from extraction.document_processor import DocumentProcessor

# With Ollama
processor = DocumentProcessor(
    input_directory="path/to/documents",     # Required
    output_format="json",                    # "json", "xml", or "csv"
    use_ai=True,                             # Enable AI analysis
    ai_provider="ollama",                    # AI provider to use
    ai_model="llama3",                       # AI model to use
    api_base="http://localhost:11434/api",   # API URL (provider-specific)
    temperature=0.1,                         # Temperature for generation
    max_tokens=2000,                         # Max tokens for response
    timeout=60,                              # API timeout in seconds
    ai_features="all",                       # AI features to enable
    ai_cache_enabled=True,                   # Enable response caching
    ai_cache_size=100                        # Number of responses to cache
)

# With OpenAI
processor = DocumentProcessor(
    input_directory="path/to/documents",
    output_format="json",
    use_ai=True,
    ai_provider="openai",
    ai_model="gpt-4o",
    api_key="your-api-key",
    temperature=0.1,
    max_tokens=2000,
    ai_features="summary,topics,suggestions"
)

# With Anthropic
processor = DocumentProcessor(
    input_directory="path/to/documents",
    output_format="json",
    use_ai=True,
    ai_provider="anthropic",
    ai_model="claude-3-opus-20240229",
    api_key="your-api-key",
    ai_features="all"
)
```

#### Methods

##### `find_documents() -> List[Path]`
Finds all Word documents in the input directory.

##### `process_document(file_path: Path) -> ProcessedDocument`
Processes a single document.

##### `process_all() -> List[ProcessedDocument]`
Processes all documents in the input directory.

##### `save_results(output_path: str) -> None`
Saves the processed results to the specified path.

### Data Classes

#### DocumentMetadata
```python
@dataclass
class DocumentMetadata:
    title: str
    author: str
    created_date: str
    last_modified_date: str
    file_path: str
    filename: str
```

#### DocumentSection
```python
@dataclass
class DocumentSection:
    title: str
    level: int
    content: str
    entities: List[Dict[str, Any]]
```

#### AiAnalysis
```python
@dataclass
class AiAnalysis:
    summary: str
    key_topics: List[str]
    content_categories: List[str]
    sentiment: str
    entity_relationships: List[Dict[str, Any]]
    document_quality_score: float
    improvement_suggestions: List[str]
    themes: List[Dict[str, Any]]
    key_insights: List[str]
    provider: AIProvider  # The AI provider used (OLLAMA, OPENAI, ANTHROPIC)
    model: str            # The specific model used
    processing_time: float  # Time taken for AI analysis (seconds)
```

#### ProcessedDocument
```python
@dataclass
class ProcessedDocument:
    metadata: DocumentMetadata
    sections: List[DocumentSection]
    raw_text: str
    ai_analysis: Optional[AiAnalysis]
    
    def to_dict(self) -> Dict[str, Any]:
        # Converts the document to a dictionary
```

## Troubleshooting

### AI Integration Issues

#### Ollama Issues

##### Connection Error
```
Failed to get AI analysis: 404, Connection refused
```

**Solution**: Ensure Ollama is running:
```bash
ollama serve
```

##### Model Not Found
```
Failed to get AI analysis: 404, model 'model_name' not found
```

**Solution**: Pull the requested model:
```bash
ollama pull model_name
```

#### OpenAI Issues

##### Authentication Error
```
OpenAI API error: 401, {"error":{"message":"Invalid Authentication"}}
```

**Solution**: Check your API key is valid and properly set:
```bash
extraction --input-dir /path/to/docs --use-ai --ai-provider openai --api-key "your-correct-api-key"
```

#### Anthropic Issues

##### Authentication Error
```
Anthropic API error: 401, {"error":"unauthorized"}
```

**Solution**: Check your API key is valid and properly set:
```bash
extraction --input-dir /path/to/docs --use-ai --ai-provider anthropic --api-key "your-correct-api-key"
```

#### General Issues

##### AI Analysis Timeout
```
Error during AI analysis: HTTPSConnectionPool timeout
```

**Solution**: Increase the timeout:
```bash
extraction --input-dir /path/to/docs --use-ai --ai-timeout 120
```

##### Response Parsing Error
```
Failed to parse AI response as JSON
```

**Solution**: Try a different model or provider that produces more consistent JSON outputs:
```bash
extraction --input-dir /path/to/docs --use-ai --ai-provider openai  # OpenAI is more reliable for JSON output
```

### Document Processing Issues

#### Section Identification Problems
If your document's sections aren't properly identified, ensure your Word document uses standard heading styles (Heading 1, Heading 2, etc.).

#### Large Document Handling
For very large documents (100+ pages), you may get better results by splitting them into smaller documents before processing.

## Best Practices

### Document Preparation
- Use standard Word heading styles (Heading 1, Heading 2) for proper section identification
- Include proper document metadata (title, author) in Word properties
- For best results with AI analysis, ensure document text is well-structured

### Processing Efficiency
- Process related documents together in the same directory
- Use JSON format for most comprehensive results
- CSV format is most convenient for spreadsheet analysis
- For large document collections, process in batches

### AI Analysis
- Use specific AI features when you don't need all analysis types
- Choose the right provider for your needs:
  - Ollama for local processing and privacy
  - OpenAI for reliability and high-quality results
  - Anthropic for nuanced document understanding
- For quality scoring and improvement suggestions, use lower temperature settings
- For creative insights and theme analysis, higher temperature may be beneficial
- Enable caching for better performance when processing similar documents
- Use different models for different document types (technical, creative, business)

---

For further assistance, please file an issue on our GitHub repository or contact support.