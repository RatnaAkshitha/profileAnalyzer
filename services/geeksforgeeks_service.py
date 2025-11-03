import requests
from urllib.parse import urlparse
from typing import Optional, Dict, Any

HEADERS = {"User-Agent": "Mozilla/5.0"}

def extract_username_from_url(profile_url: str) -> Optional[str]:
    """
    Extracts GeeksforGeeks username from profile URL.
    Example:
        https://auth.geeksforgeeks.org/user/johndoe/practice/ -> johndoe
    """
    try:
        path = urlparse(profile_url).path.strip("/")
        parts = path.split("/")
        if "user" in parts:
            idx = parts.index("user")
            return parts[idx + 1] if idx + 1 < len(parts) else None
    except Exception:
        return None
    return None


def get_geeksforgeeks_data(profile_url: str) -> Dict[str, Any]:
    """
    Fetches GFG user profile data via API.
    Returns dictionary with coding score, problems solved, and rank.
    """
    if not profile_url:
        return {"skipped": True, "message": "Profile URL not provided"}

    username = extract_username_from_url(profile_url)
    if not username:
        return {
            "error": "Invalid profile URL",
            "expected_format": "https://auth.geeksforgeeks.org/user/<username>/practice/"
        }

    api_url = f"https://practiceapi.geeksforgeeks.org/api/v1/user/{username}/"

    try:
        response = requests.get(api_url, headers=HEADERS, timeout=10)

        if response.status_code == 404:
            return {"error": "Profile not found on GeeksforGeeks"}

        response.raise_for_status()
        data = response.json()

        return {
            "username": username,
            "profile_url": profile_url,
            "coding_score": data.get("coding_score", "N/A"),
            "problems_solved": data.get("total_problems_solved", "N/A"),
            "institute_rank": data.get("institute_rank", "N/A"),
        }

    except requests.Timeout:
        return {"error": "Request timed out"}
    except requests.RequestException as e:
        return {"error": f"Network error: {str(e)}"}
    except ValueError:
        return {"error": "Invalid JSON response from GeeksforGeeks API"}
