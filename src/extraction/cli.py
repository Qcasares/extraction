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

    parser.add_argument(
        "--use-ai",
        action="store_true",
        help="Use Ollama AI for enhanced analysis",
    )

    parser.add_argument(
        "--model",
        "-m",
        type=str,
        default="llama3",
        help="Ollama model to use (default: llama3)",
    )

    parser.add_argument(
        "--api-base",
        type=str,
        default="http://localhost:11434/api",
        help="Base URL for Ollama API (default: http://localhost:11434/api)",
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
        # Create processor
        processor = DocumentProcessor(
            input_directory=parsed_args.input_dir,
            output_format=parsed_args.format,
            use_ai=parsed_args.use_ai,
            ollama_model=parsed_args.model if parsed_args.use_ai else None,
            ollama_api_base=parsed_args.api_base,
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
