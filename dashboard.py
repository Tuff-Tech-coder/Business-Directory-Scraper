"""Loopback-only, offline demonstration UI using the standard library."""

import argparse
import io
import json
import logging
from dataclasses import asdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from directory_scraper import export_to_excel, filename_slug, run_demo

STATIC = Path(__file__).parent / "web"
CITIES = {"Austin, TX", "Chicago, IL", "Seattle, WA"}
CATEGORIES = {"plumbers", "dentists", "electricians"}


def demo_request(query):
    params = parse_qs(query, keep_blank_values=True)
    category = params.get("category", ["plumbers"])[0]
    city = params.get("city", ["Austin, TX"])[0]
    count = int(params.get("count", ["35"])[0])
    seed = int(params.get("seed", ["42"])[0])
    if category not in CATEGORIES or city not in CITIES or not 1 <= count <= 1000:
        raise ValueError("Choose a supported category, city, and 1–1000 records")
    results, stats = run_demo(category, city, count, seed)
    return results, stats, params


def filtered_records(results, params):
    search = params.get("search", [""])[0].casefold().strip()
    rating = float(params.get("rating", ["0"])[0])
    if rating not in {0, 3, 4, 4.5, 5}:
        raise ValueError("Unsupported rating filter")
    website = params.get("website", ["false"])[0] == "true"
    return [
        b
        for b in results
        if (
            not search
            or search in " ".join((b.name, b.address, b.city_state, b.category)).casefold()
        )
        and (not rating or (b.rating and float(b.rating.split("/")[0]) >= rating))
        and (not website or b.website)
    ]


class DemoHandler(BaseHTTPRequestHandler):
    def send_content(self, status, content, mime, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; frame-ancestors 'none'",
        )
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        # Do not serve this local tool through arbitrary Host headers.
        host = self.headers.get("Host", "").split(":")[0]
        if host not in {"127.0.0.1", "localhost"}:
            self.send_content(403, b"Local access only", "text/plain")
            return
        path = urlsplit(self.path)
        files = {
            "/": ("index.html", "text/html; charset=utf-8"),
            "/app.css": ("app.css", "text/css"),
            "/app.js": ("app.js", "text/javascript"),
        }
        if path.path in files:
            filename, mime = files[path.path]
            self.send_content(200, (STATIC / filename).read_bytes(), mime)
            return
        if path.path not in {"/api/demo", "/api/export"}:
            self.send_content(404, b"Not found", "text/plain")
            return
        try:
            results, stats, params = demo_request(path.query)
            if path.path == "/api/demo":
                payload = {"businesses": [asdict(b) for b in results], "stats": stats}
                self.send_content(200, json.dumps(payload).encode(), "application/json")
            else:
                results = filtered_records(results, params)
                if not results:
                    raise ValueError("No matching records to export")
                buffer = io.BytesIO()
                export_to_excel(results, buffer, source=stats["source"])
                filename = f"{filename_slug(params.get('category', ['plumbers'])[0])}_demo.xlsx"
                self.send_content(
                    200,
                    buffer.getvalue(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename,
                )
        except (ValueError, OverflowError) as exc:
            self.send_content(400, json.dumps({"error": str(exc)}).encode(), "application/json")
        except Exception:
            logging.exception("Demo request failed")
            self.send_content(
                500,
                b'{"error":"The demo could not finish. Check the server log."}',
                "application/json",
            )

    def log_message(self, format, *args):
        logging.getLogger(__name__).info(format, *args)


def serve(port=8765):
    with ThreadingHTTPServer(("127.0.0.1", port), DemoHandler) as server:
        print(f"Business Directory dashboard: http://127.0.0.1:{server.server_port}", flush=True)
        print("Offline demo. Press Ctrl+C to stop.", flush=True)
        server.serve_forever()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=8765)
    serve(parser.parse_args().port)
