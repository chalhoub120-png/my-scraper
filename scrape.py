import os
import re
import time
import hashlib
import requests

os.makedirs("content", exist_ok=True)
LOG = []
API = "https://api.firecrawl.dev/v1"


def log(msg):
    print(msg)
    LOG.append(str(msg))


def slug(t):
    s = re.sub(r"[^a-zA-Z0-9]+", "-", t).strip("-")[:60]
    return s or "q-" + hashlib.md5(t.encode()).hexdigest()[:8]


def save(name, text):
    with open(f"content/{name}.md", "w", encoding="utf-8") as f:
        f.write(text)
    log("OK saved: " + name)


def post(path, body, key):
    r = requests.post(
        API + path,
        headers={"Authorization": "Bearer " + key},
        json=body,
        timeout=180,
    )
    if not r.ok:
        log(f"ERROR {r.status_code}: {r.text[:300]}")
        return None
    return r.json()


def get(path_or_url, key):
    url = path_or_url if path_or_url.startswith("http") else API + path_or_url
    r = requests.get(
        url,
        headers={"Authorization": "Bearer " + key},
        timeout=180,
    )
    if not r.ok:
        log(f"ERROR {r.status_code}: {r.text[:300]}")
        return None
    return r.json()


def crawl(url, key, limit=10):
    """Start a crawl, wait for it, and return all pages (follows pagination)."""
    d = post("/crawl", {"url": url, "limit": limit,
             "scrapeOptions": {"formats": ["markdown"]}}, key)
    if not d or not d.get("id"):
        return None
    job_id = d["id"]
    log("Crawl started: " + job_id)
    max_wait = max(300, limit * 10)
    waited = 0
    while waited < max_wait:
        time.sleep(5)
        waited += 5
        s = get("/crawl/" + job_id, key)
        if not s:
            return None
        status = s.get("status")
        if status == "completed":
            pages = list(s.get("data", []))
            nxt = s.get("next")
            while nxt:
                s = get(nxt, key)
                if not s:
                    break
                pages += s.get("data", [])
                nxt = s.get("next")
            return pages
        if status in ("failed", "cancelled"):
            log("ERROR: crawl " + status)
            return None
    log("ERROR: crawl timed out")
    return None


def save_crawl(url, pages):
    """One file per page, inside content/<crawl-name>/."""
    folder = "crawl-" + slug(url)
    os.makedirs(f"content/{folder}", exist_ok=True)
    for i, p in enumerate(pages, 1):
        meta = p.get("metadata", {})
        src = meta.get("sourceURL", "")
        title = meta.get("title", "")
        name = f"{folder}/{i:04d}-{slug(src)[:50]}"
        save(name, f"# {title}\nSource: {src}\n\n{p.get('markdown', '')}\n")
    log(f"OK crawl pages: {len(pages)}")


def main():
    key = os.environ.get("FIRECRAWL_API_KEY", "").strip()
    if not key:
        log("ERROR: the secret FIRECRAWL_API_KEY is empty or missing")
        return
    if not os.path.exists("urls.txt"):
        log("ERROR: urls.txt not found")
        return
    with open("urls.txt", encoding="utf-8") as f:
        lines = [l.strip() for l in f if l.strip() and not l.startswith("#")]
    if not lines:
        log("ERROR: urls.txt is empty")
        return
    for line in lines:
        log("Working on: " + line)
        try:
            if line.lower().startswith("search:"):
                q = line[7:].strip()
                d = post("/search", {"query": q, "limit": 8,
                         "scrapeOptions": {"formats": ["markdown"]}}, key)
                if d:
                    parts = [
                        f"## {i}. {x.get('title', '')}\nSource: {x.get('url', '')}\n\n"
                        f"{x.get('markdown') or x.get('description', '')}\n"
                        for i, x in enumerate(d.get("data", []), 1)
                    ]
                    save("search-" + slug(q), "\n---\n\n".join(parts))
            elif line.lower().startswith("crawl:"):
                # format:  crawl: https://site.com
                #     or:  crawl: https://site.com | 200
                body = line[6:].strip()
                limit = 10
                if "|" in body:
                    body, lim = body.rsplit("|", 1)
                    body = body.strip()
                    if lim.strip().isdigit():
                        limit = int(lim.strip())
                pages = crawl(body, key, limit)
                if pages is not None:
                    save_crawl(body, pages)
            else:
                d = post("/scrape", {"url": line, "formats": ["markdown"]}, key)
                if d:
                    save(slug(line), d["data"].get("markdown", ""))
        except Exception as e:
            log("ERROR: " + repr(e))


try:
    main()
except Exception as e:
    log("FATAL: " + repr(e))

with open("content/log.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(LOG))
