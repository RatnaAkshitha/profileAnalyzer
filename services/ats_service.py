import os
import requests
import random

ATS_API_URL = os.getenv("ATS_API_URL")  # POST {"resume_url": "<url>"}
ATS_API_KEY = os.getenv("ATS_API_KEY")

def get_ats_score(resume_url: str):
    if not resume_url or resume_url.strip().lower() == "n/a":
        # Skip scoring if user entered "N/A" or empty
        return {"score": None, "remarks": "No ATS scoring (resume not provided)."}
    
    if ATS_API_URL:
        try:
            headers = {"Content-Type": "application/json"}
            if ATS_API_KEY:
                headers["Authorization"] = f"Bearer {ATS_API_KEY}"
            payload = {"resume_url": resume_url}
            resp = requests.post(ATS_API_URL, json=payload, headers=headers, timeout=20)
            resp.raise_for_status()
            data = resp.json()
            # Expecting a structure like {"score": 78, "remarks": "..."}
            score = data.get("score") or data.get("ats_score") or None
            remarks = data.get("remarks") or data.get("message") or ""
            return {"score": score, "remarks": remarks, "raw_response": data}
        except requests.RequestException as e:
            return {"error": f"ATS API request failed: {str(e)}"}
        except Exception as e:
            return {"error": f"ATS API error: {str(e)}"}
    else:
        # Mock scoring when no ATS endpoint is provided
        score = random.randint(60, 95)
        remarks = "Mock ATS score (provide ATS_API_URL in .env for real scoring)."
        return {"score": score, "remarks": remarks}
