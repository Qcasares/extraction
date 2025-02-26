"""Tests for CLI module."""

from unittest.mock import MagicMock, patch

import pytest

from extraction.cli import main, parse_args


def test_parse_args() -> None:
    """Test argument parsing."""
    args = parse_args(
        [
            "--input-dir",
            "/test/input",
            "--output-dir",
            "/test/output",
            "--format",
            "xml",
            "--use-ai",
            "--model",
            "llama3",
            "--verbose",
        ]
    )

    assert args.input_dir == "/test/input"
    assert args.output_dir == "/test/output"
    assert args.format == "xml"
    assert args.use_ai is True
    assert args.model == "llama3"
    assert args.verbose is True


def test_parse_args_defaults() -> None:
    """Test argument parsing with defaults."""
    args = parse_args(["--input-dir", "/test/input"])

    assert args.input_dir == "/test/input"
    assert args.output_dir == "./output"
    assert args.format == "json"
    assert args.use_ai is False
    assert args.model == "llama3"
    assert args.verbose is False


@patch("extraction.cli.DocumentProcessor")
def test_main_success(mock_processor_class: MagicMock) -> None:
    """Test main function with successful execution."""
    # Mock the processor
    mock_processor = MagicMock()
    mock_processor_class.return_value = mock_processor

    # Run main
    exit_code = main(["--input-dir", "/test/input"])

    # Verify processor was created with correct args
    mock_processor_class.assert_called_once_with(
        input_directory="/test/input",
        output_format="json",
        use_ai=False,
        ollama_model=None,
        ollama_api_base="http://localhost:11434/api",
    )

    # Verify methods were called
    mock_processor.process_all.assert_called_once()
    mock_processor.save_results.assert_called_once_with("./output")

    # Verify exit code
    assert exit_code == 0


@patch("extraction.cli.DocumentProcessor")
def test_main_with_ai(mock_processor_class: MagicMock) -> None:
    """Test main function with AI enabled."""
    # Mock the processor
    mock_processor = MagicMock()
    mock_processor_class.return_value = mock_processor

    # Run main with AI
    exit_code = main(["--input-dir", "/test/input", "--use-ai", "--model", "llama3"])

    # Verify processor was created with correct args
    mock_processor_class.assert_called_once_with(
        input_directory="/test/input",
        output_format="json",
        use_ai=True,
        ollama_model="llama3",
        ollama_api_base="http://localhost:11434/api",
    )

    # Verify exit code
    assert exit_code == 0


@patch("extraction.cli.DocumentProcessor")
def test_main_error(mock_processor_class: MagicMock) -> None:
    """Test main function with error."""
    # Make the processor raise an exception
    mock_processor_class.side_effect = ValueError("Test error")

    # Run main
    exit_code = main(["--input-dir", "/test/input"])

    # Verify exit code indicates error
    assert exit_code == 1
