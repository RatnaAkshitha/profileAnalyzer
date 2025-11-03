
import requests
import time

REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/json",
    "Origin": "https://leetcode.com",
    "Referer": "https://leetcode.com/",
}

def get_leetcode_data(username: str, retries: int = 3):
    if not username:
        return {"skipped": True}
    query = """
    query getUserProfile($username: String!) {
        matchedUser(username: $username) {
            username
            profile { ranking }
            languageProblemCount { languageName problemsSolved }
            submitStats: submitStatsGlobal {
                acSubmissionNum { difficulty count }
            }
            badges { name }
        }
    }
    """
    for attempt in range(retries):
        try:
            resp = requests.post(
                "https://leetcode.com/graphql",
                json={"query": query, "variables": {"username": username}},
                headers=REQUEST_HEADERS,
                timeout=10
            )
            resp.raise_for_status()
            data = resp.json()
            if data.get("errors"):
                return {"error": data["errors"][0].get("message", "API error")}
            user = data.get("data", {}).get("matchedUser")
            if not user:
                return {"error": "User not found or profile private"}

            # parse stats
            problems = {}
            total = 0
            submit_stats = user.get("submitStats") or {}
            ac = submit_stats.get("acSubmissionNum") or []
            for stat in ac:
                if stat.get("difficulty") == "All":
                    total = stat.get("count", 0)
                else:
                    problems[stat.get("difficulty")] = stat.get("count", 0)

            languages = {}
            for lang in user.get("languageProblemCount") or []:
                languages[lang.get("languageName")] = lang.get("problemsSolved")

            badges = [b.get("name") for b in (user.get("badges") or [])]

            return {
                "username": username,
                "rank": user.get("profile", {}).get("ranking"),
                "total_solved": total,
                "problems_by_difficulty": problems,
                "languages": languages,
                "badges": badges
            }
        except requests.RequestException as e:
            if attempt == retries - 1:
                return {"error": f"Network error: {str(e)}"}
            time.sleep(1)
    return {"error": "Failed after retries"}