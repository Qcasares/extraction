"""Data extraction module."""

from typing import Any, Dict, List, Optional

import requests


class Extractor:
    """Base class for data extractors."""

    def __init__(self, base_url: str, timeout: int = 30) -> None:
        """Initialize the extractor.

        Args:
            base_url: The base URL for the API.
            timeout: Timeout in seconds for requests.
        """
        self.base_url = base_url
        self.timeout = timeout
        self.session = requests.Session()

    def extract(
        self, endpoint: str, params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Extract data from the API.

        Args:
            endpoint: The API endpoint to hit.
            params: Query parameters to include in the request.

        Returns:
            The parsed JSON response.

        Raises:
            requests.exceptions.RequestException: If the request fails.
        """
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        response = self.session.get(url, params=params, timeout=self.timeout)
        response.raise_for_status()
        result: Dict[str, Any] = response.json()
        return result
