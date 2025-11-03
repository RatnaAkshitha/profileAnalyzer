import requests
import os
from collections import defaultdict
from datetime import datetime
from dateutil import parser as dateparser

GITHUB_API = "https://api.github.com"
TOKEN = os.getenv("GITHUB_TOKEN")  # optional; increases rate limits

HEADERS = {"User-Agent": "Profile-Analyzer"}
if TOKEN:
    HEADERS["Authorization"] = f"token {TOKEN}"

def get_github_data(username: str):
    if not username:
        return {"skipped": True}
    try:
        user_url = f"{GITHUB_API}/users/{username}"
        r = requests.get(user_url, headers=HEADERS, timeout=10)
        if r.status_code == 404:
            return {"error": "GitHub user not found"}
        r.raise_for_status()
        user_info = r.json()

        # total public repos
        total_repos = user_info.get("public_repos", 0)

        # estimate commits/activity: fetch recent public events (max 300 events via pagination)
        events_url = f"{GITHUB_API}/users/{username}/events/public"
        events = []
        page = 1
        while page <= 3:  # up to 3 pages (30 per page default) -> ~90 events
            er = requests.get(events_url, headers=HEADERS, params={"page": page}, timeout=10)
            if er.status_code != 200:
                break
            page_events = er.json()
            if not page_events:
                break
            events.extend(page_events)
            page += 1

        # Count PushEvent per month to estimate active months
        push_events = [e for e in events if e.get("type") == "PushEvent"]
        months = set()
        for ev in push_events:
            created = ev.get("created_at")
            if created:
                dt = dateparser.parse(created)
                months.add((dt.year, dt.month))

        active_months = len(months)
        recent_pushes = len(push_events)

        return {
            "username": username,
            "total_repos": total_repos,
            "recent_push_events": recent_pushes,
            "active_months": active_months,
            "profile_url": user_info.get("html_url")
        }
    except requests.RequestException as e:
        return {"error": f"Network error: {str(e)}"}
    except Exception as e:
        return {"error": f"Error: {str(e)}"}
