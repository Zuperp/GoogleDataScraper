# google_scraper.py

import time
import os

def search_keyword(query, api_key, google_domain="google.dk", hl="da", gl="dk", mock_mode=False, mock_value=None, api_delay=0.5):
    """
    Returns (hits, error_message) tuple.
    If mock_mode, uses mock_value and returns (int, None).
    If real mode, calls SerpAPI and returns (int, None) or (None, error_message) on error.
    """
    if mock_mode:
        # Simulate network delay for realism
        time.sleep(api_delay)
        if mock_value is not None:
            try:
                return int(mock_value), None
            except Exception:
                return None, "Invalid mock value"
        else:
            return None, "No mock value provided"

    # Real API mode
    try:
        from serpapi import GoogleSearch
    except ImportError:
        return None, "serpapi module not installed"

    if not api_key:
        return None, "Missing SerpAPI API key"

    params = {
        "engine": "google",
        "q": query,
        "api_key": api_key,
        "google_domain": google_domain,
        "hl": hl,
        "gl": gl
    }

    try:
        search = GoogleSearch(params)
        result = search.get_dict()
        info = result.get("search_information", {})
        hits = info.get("total_results")
        if hits is None:
            return None, "No 'total_results' found"
        return hits, None
    except Exception as e:
        return None, f"API error: {e}"
