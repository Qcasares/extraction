"""Command-line interface for the document processor."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from extraction.document_processor import DocumentProcessor


def parse_args(args: Optional[List[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments for the Word document processor.

    Args:
        args: Command-line arguments (uses sys.argv if None)

    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Process Word documents and extract structured content"
    )

    parser.add_argument(
        "--input-dir",
        "-i",
        type=str,
        required=True,
        help="Directory containing Word documents to process",
    )

    parser.add_argument(
        "--output-dir",
        "-o",
        type=str,
        default="./output",
        help="Directory to save processed results (default: ./output)",
    )

    parser.add_argument(
        "--format",
        "-f",
        type=str,
        choices=["json", "xml", "csv"],
        default="json",
        help="Output format (default: json)",
    )

    # AI analysis options
    ai_group = parser.add_argument_group("AI Analysis Options")
    
    ai_group.add_argument(
        "--use-ai",
        action="store_true",
        help="Enable AI analysis",
    )
    
    ai_group.add_argument(
        "--ai-provider",
        type=str,
        choices=["ollama", "openai", "anthropic", "none"],
        default="ollama",
        help="AI provider to use (default: ollama)",
    )

    ai_group.add_argument(
        "--model",
        "-m",
        type=str,
        help=("AI model to use (provider-specific, defaults: "
              "ollama: 'llama3', openai: 'gpt-4o' (also supports gpt-4o-mini, o1-preview, o1-mini, o3-preview, o3-mini), "
              "anthropic: 'claude-3-opus-20240229')"),
    )

    ai_group.add_argument(
        "--api-base",
        type=str,
        help=("Base URL for AI provider API (defaults: "
              "ollama: 'http://localhost:11434/api', openai: 'https://api.openai.com/v1', "
              "anthropic: 'https://api.anthropic.com/v1')"),
    )
    
    ai_group.add_argument(
        "--api-key",
        type=str,
        default="",
        help="API key for authentication (required for OpenAI and Anthropic)",
    )
    
    ai_group.add_argument(
        "--temperature",
        type=float,
        default=0.1,
        help="Temperature parameter for AI generation (0.0-1.0, lower for more focused outputs). Note: Not supported for OpenAI o-series models (gpt-4o, o1-*, o3-*). (default: 0.1)",
    )
    
    ai_group.add_argument(
        "--max-tokens",
        type=int,
        default=2000,
        help="Maximum number of tokens for AI response (used as max_completion_tokens for OpenAI and Anthropic models) (default: 2000)",
    )
    
    ai_group.add_argument(
        "--ai-timeout",
        type=int,
        default=60,
        help="Timeout in seconds for AI API calls (default: 60)",
    )
    
    ai_group.add_argument(
        "--ai-features",
        type=str,
        default="all",
        help=("Comma-separated list of AI analysis features to enable: "
              "summary,topics,categories,sentiment,relationships,quality,suggestions,themes,insights "
              "(can also be 'all' or 'none', default: all)"),
    )
    
    ai_group.add_argument(
        "--disable-ai-cache",
        action="store_true",
        help="Disable caching of AI responses (cache is enabled by default)",
    )
    
    ai_group.add_argument(
        "--ai-cache-size",
        type=int,
        default=100,
        help="Maximum number of AI responses to cache (default: 100)",
    )

    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable verbose logging",
    )

    return parser.parse_args(args)


def main(args: Optional[List[str]] = None) -> int:
    """Run the document processor.

    Args:
        args: Command-line arguments (uses sys.argv if None)

    Returns:
        Exit code
    """
    parsed_args = parse_args(args)

    # Configure logging
    log_level = logging.DEBUG if parsed_args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger(__name__)

    try:
        # Create processor with enhanced AI configuration
        processor = DocumentProcessor(
            input_directory=parsed_args.input_dir,
            output_format=parsed_args.format,
            use_ai=parsed_args.use_ai,
            ai_provider=parsed_args.ai_provider,
            ai_model=parsed_args.model,
            api_base=parsed_args.api_base,
            api_key=parsed_args.api_key,
            temperature=parsed_args.temperature,
            max_tokens=parsed_args.max_tokens,
            timeout=parsed_args.ai_timeout,
            ai_features=parsed_args.ai_features,
            ai_cache_enabled=not parsed_args.disable_ai_cache,
            ai_cache_size=parsed_args.ai_cache_size,
        )

        # Process documents
        logger.info("Starting document processing")
        processor.process_all()

        # Save results
        logger.info("Saving results to %s", parsed_args.output_dir)
        processor.save_results(parsed_args.output_dir)

        logger.info("Processing complete")
        return 0

    except Exception as e:
        logger.error("Error during processing: %s", str(e), exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
