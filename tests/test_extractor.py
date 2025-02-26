"""Tests for the extractor module."""

from unittest.mock import MagicMock, patch

import pytest
import requests

from extraction.extractor import Extractor


@pytest.fixture
def extractor() -> Extractor:
    """Create an extractor instance for testing."""
    return Extractor(base_url="https://api.example.com")


def test_extract_successful(extractor: Extractor) -> None:
    """Test successful data extraction."""
    # Mock the response
    mock_response = MagicMock()
    mock_response.json.return_value = {"data": "test_data"}

    # Patch the session.get method
    with patch.object(extractor.session, "get", return_value=mock_response) as mock_get:
        result = extractor.extract("test/endpoint", params={"param": "value"})

        # Verify the request was made correctly
        mock_get.assert_called_once_with(
            "https://api.example.com/test/endpoint",
            params={"param": "value"},
            timeout=30,
        )

        # Verify the response was processed correctly
        assert result == {"data": "test_data"}


def test_extract_failure(extractor: Extractor) -> None:
    """Test extraction failure handling."""
    # Mock a failed response
    mock_response = MagicMock()
    mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError(
        "404 Not Found"
    )

    # Patch the session.get method
    with patch.object(extractor.session, "get", return_value=mock_response) as mock_get:
        # Verify the exception is propagated
        with pytest.raises(requests.exceptions.HTTPError):
            extractor.extract("test/endpoint")

        # Verify the request was made
        mock_get.assert_called_once()
