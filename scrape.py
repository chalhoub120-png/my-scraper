import os
import re
import sys
import requests

API_KEY = os.environ.get("FIRECRAWL_API_KEY")
if not API_KEY:
    sys.exit("Missing FIRECRAWL_API_KEY secret")

os.makedirs("content", exist_ok=True)

with open("urls.txt") as f:
    urls = [l.strip() for l in f if l.strip() and not l.startswith("#")]

for url in urls:
    print("Scraping:", url)
    r = requests.post(
        "https://api.firecrawl.dev/v1/scrape",
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json",
        },
        json={"url": url, "formats": ["markdown"]},
        timeout=120,
    )
    r.raise_for_status()
    markdown = r.json()["data"]["markdown"]
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", url).strip("-")[:80]
    with open(f"content/{slug}.md", "w", encoding="utf-8") as out:
        out.write(markdown)
    print("Saved:", slug)
