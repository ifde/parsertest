import asyncio
import io
import mimetypes
import re
from pathlib import Path
from urllib.parse import urlparse, urljoin

import requests
from flask import Flask, request, send_file, render_template_string, abort
from playwright.async_api import async_playwright

PROJECT_ROOT = Path(__file__).parent
app = Flask(__name__)

HTML = """
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>Simple Image Downloader</title></head>
  <body style="font-family: sans-serif; max-width: 720px; margin: 2rem auto;">
    <h1>Simple Image Downloader</h1>
    <form method="post" action="/download" style="display:flex; gap:.5rem;">
      <input type="url" name="url" placeholder="Paste page or image URL" required style="flex:1; padding:.5rem;">
      <button type="submit" style="padding:.5rem 1rem;">Download image</button>
    </form>
    <p style="color:#555; margin-top:1rem;">Paste a direct image link or a page containing an image (og:image or the first &lt;img&gt;).</p>
  </body>
</html>
"""

IMAGE_EXT_RE = re.compile(r"\.(jpg|jpeg|png|gif|webp|bmp|tiff)$", re.IGNORECASE)

def is_likely_image_url(url: str) -> bool:
    return bool(IMAGE_EXT_RE.search(urlparse(url).path))

def filename_from_url(url: str, content_type: str | None) -> str:
    name = Path(urlparse(url).path).name or "download"
    if "." not in name and content_type:
        ext = mimetypes.guess_extension(content_type.split(";")[0].strip()) or ".bin"
        name += ext
    return name

def absolutize_url(u: str, base: str) -> str:
    if not u:
        return u
    if u.startswith("//"):
        return "https:" + u
    if u.startswith("/"):
        return urljoin(base, u)
    return u

def fetch_image_bytes(url: str, referer: str | None = None) -> tuple[bytes, str, str]:
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    }
    if referer:
        headers["Referer"] = referer
    
    resp = requests.get(url, headers=headers, timeout=25)
    if resp.status_code != 200:
        abort(400, description=f"Failed to fetch image: HTTP {resp.status_code}")
    content_type = resp.headers.get("Content-Type", "")
    data = resp.content
    fname = filename_from_url(url, content_type)
    return data, content_type, fname

async def extract_image_url(page_url: str) -> str | None:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        try:
            page = await browser.new_page()
            await page.goto(page_url, wait_until="domcontentloaded", timeout=25000)

            # Prefer og:image
            og = await page.locator('meta[property="og:image"]').first.get_attribute("content")
            if og and og.strip():
                return absolutize_url(og.strip(), page_url)

            # Fallback to first non-data <img>
            img_urls = await page.evaluate(
                "() => Array.from(document.images)"
                ".map(img => img.src)"
                ".filter(u => u && !u.startsWith('data:'))"
            )
            if img_urls:
                return absolutize_url(img_urls[0], page_url)
            return None
        finally:
            await browser.close()

@app.get("/")
def index():
    return render_template_string(HTML)

@app.post("/download")
def download():
    url = request.form.get("url", "").strip()
    if not url:
        abort(400, description="URL is required.")

    # If it looks like a direct image, fetch it
    if is_likely_image_url(url):
        data, content_type, fname = fetch_image_bytes(url)
        return send_file(io.BytesIO(data), mimetype=content_type or "application/octet-stream",
                         as_attachment=True, download_name=fname)

    # Otherwise try to extract an image from the page
    img_url = asyncio.run(extract_image_url(url))
    if not img_url:
        abort(404, description="No image found on the page.")
    data, content_type, fname = fetch_image_bytes(img_url, referer=url)
    return send_file(io.BytesIO(data), mimetype=content_type or "application/octet-stream",
                     as_attachment=True, download_name=fname)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)