import ipaddress
import logging
import os
import socket
import threading
from urllib.parse import urljoin, urlparse

import requests
from flask import Flask, Response, jsonify, request, send_from_directory
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_FILE = os.getenv("INDEX_FILE", "sec-data-app.html")
HOST = os.getenv("PROXY_HOST", "127.0.0.1")
PORT = int(os.getenv("PROXY_PORT", "8000"))

CONNECT_TIMEOUT = float(os.getenv("PROXY_CONNECT_TIMEOUT", "5"))
READ_TIMEOUT = float(os.getenv("PROXY_READ_TIMEOUT", "30"))
MAX_REDIRECTS = int(os.getenv("PROXY_MAX_REDIRECTS", "5"))
MAX_URL_LENGTH = int(os.getenv("PROXY_MAX_URL_LENGTH", "12000"))

# Default API hosts used by the SEC dashboard. Add more through:
#   set PROXY_ALLOWED_HOSTS=api.example.com,another.example.com
DEFAULT_ALLOWED_HOSTS = {
    "api.tiingo.com",
    "api.stlouisfed.org",
    "gamma-api.polymarket.com",
    "api.rss2json.com",
    "api.usaspending.gov",
    "www.sec.gov",
    "feeds.finance.yahoo.com",
    "news.google.com",
    "finnhub.io",
    "mcp.exa.ai",
}

extra_hosts = {
    h.strip().lower()
    for h in os.getenv("PROXY_ALLOWED_HOSTS", "").split(",")
    if h.strip()
}
ALLOWED_HOSTS = DEFAULT_ALLOWED_HOSTS | extra_hosts

# Only these local browser origins receive CORS permission.
ALLOWED_ORIGINS = {
    f"http://localhost:{PORT}",
    f"http://127.0.0.1:{PORT}",
}

app = Flask(__name__, static_folder=BASE_DIR, static_url_path="")

logging.basicConfig(
    level=os.getenv("PROXY_LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("dashboard-proxy")

_thread_local = threading.local()


# ──────────────────────────────────────────────────────────────────────────────
# HTTP client
# ──────────────────────────────────────────────────────────────────────────────

def _make_session() -> requests.Session:
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        other=0,
        backoff_factor=0.6,
        status_forcelist=(408, 425, 429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
        respect_retry_after_header=True,
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=20,
        pool_maxsize=20,
    )

    session = requests.Session()
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/151.0 Safari/537.36 SEC-Dashboard-Local-Proxy/2.0"
            ),
            "Accept": "application/json,text/plain,*/*",
            "Accept-Language": "en-US,en;q=0.9",
        }
    )
    return session


def _session() -> requests.Session:
    if not hasattr(_thread_local, "session"):
        _thread_local.session = _make_session()
    return _thread_local.session


# ──────────────────────────────────────────────────────────────────────────────
# Security / URL validation
# ──────────────────────────────────────────────────────────────────────────────

def _host_is_allowed(hostname: str) -> bool:
    hostname = (hostname or "").lower().rstrip(".")
    return hostname in ALLOWED_HOSTS


def _reject_private_or_local_resolution(hostname: str) -> None:
    """
    Defense-in-depth against SSRF / DNS rebinding.
    API hosts are already allow-listed, but this also refuses hosts that resolve
    to loopback/private/link-local/reserved addresses.
    """
    try:
        infos = socket.getaddrinfo(hostname, None)
    except socket.gaierror as exc:
        raise ValueError(f"DNS lookup failed for {hostname}: {exc}") from exc

    for info in infos:
        raw_ip = info[4][0]
        try:
            ip = ipaddress.ip_address(raw_ip)
        except ValueError:
            continue

        if (
            ip.is_loopback
            or ip.is_private
            or ip.is_link_local
            or ip.is_multicast
            or ip.is_reserved
            or ip.is_unspecified
        ):
            raise ValueError(f"Refusing non-public destination address: {raw_ip}")


def _validate_target(raw_url: str) -> str:
    if not raw_url:
        raise ValueError("No URL provided")

    if len(raw_url) > MAX_URL_LENGTH:
        raise ValueError("URL is too long")

    parsed = urlparse(raw_url)

    if parsed.scheme not in {"http", "https"}:
        raise ValueError("Only http:// and https:// URLs are allowed")

    if not parsed.hostname:
        raise ValueError("Target URL has no hostname")

    if parsed.username or parsed.password:
        raise ValueError("Credentials embedded in URLs are not allowed")

    if not _host_is_allowed(parsed.hostname):
        raise ValueError(
            f"Host '{parsed.hostname}' is not allowed. "
            "Add it to PROXY_ALLOWED_HOSTS if this API is intentional."
        )

    _reject_private_or_local_resolution(parsed.hostname)
    return raw_url


# ──────────────────────────────────────────────────────────────────────────────
# Proxy helpers
# ──────────────────────────────────────────────────────────────────────────────

REQUEST_HEADERS_TO_FORWARD = {
    "accept",
    "accept-language",
    "authorization",
    "content-type",
    "if-modified-since",
    "if-none-match",
    "range",
}

RESPONSE_HEADERS_TO_FORWARD = {
    "cache-control",
    "content-disposition",
    "content-language",
    "content-range",
    "content-type",
    "etag",
    "expires",
    "last-modified",
    "retry-after",
}


def _forward_request_headers() -> dict:
    headers = {}
    for name, value in request.headers.items():
        if name.lower() in REQUEST_HEADERS_TO_FORWARD:
            headers[name] = value
    return headers


def _upstream_request(target_url: str) -> requests.Response:
    """
    Follow redirects manually so every redirect target is revalidated.
    """
    current_url = _validate_target(target_url)

    for redirect_count in range(MAX_REDIRECTS + 1):
        method = request.method

        upstream = _session().request(
            method=method,
            url=current_url,
            headers=_forward_request_headers(),
            data=request.get_data() if method not in {"GET", "HEAD"} else None,
            timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            allow_redirects=False,
        )

        if upstream.status_code not in {301, 302, 303, 307, 308}:
            return upstream

        if redirect_count >= MAX_REDIRECTS:
            upstream.close()
            raise requests.TooManyRedirects(
                f"Exceeded {MAX_REDIRECTS} upstream redirects"
            )

        location = upstream.headers.get("Location")
        if not location:
            return upstream

        next_url = urljoin(current_url, location)
        upstream.close()
        current_url = _validate_target(next_url)

    raise requests.TooManyRedirects("Redirect handling failed")


def _build_browser_response(upstream: requests.Response) -> Response:
    # requests already decompresses gzip/br content when .content is read, so
    # intentionally do NOT forward Content-Encoding or upstream Content-Length.
    payload = b"" if request.method == "HEAD" else upstream.content

    response = Response(
        payload,
        status=upstream.status_code,
    )

    for name, value in upstream.headers.items():
        if name.lower() in RESPONSE_HEADERS_TO_FORWARD:
            response.headers[name] = value

    # Debugging information visible in DevTools without exposing secrets.
    response.headers["X-Local-Proxy"] = "SEC-Dashboard-Proxy/2.0"
    return response


# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.after_request
def add_local_cors_headers(response: Response):
    origin = request.headers.get("Origin")

    if origin in ALLOWED_ORIGINS:
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Vary"] = "Origin"
        response.headers["Access-Control-Allow-Methods"] = "GET,HEAD,POST,OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = (
            "Accept,Accept-Language,Authorization,Content-Type,"
            "If-Modified-Since,If-None-Match,Range"
        )

    return response


@app.route("/health", methods=["GET"])
def health():
    return jsonify(
        {
            "ok": True,
            "service": "SEC Dashboard Local Proxy",
            "version": "2.0",
            "allowed_hosts": sorted(ALLOWED_HOSTS),
        }
    )


@app.route("/proxy", methods=["GET", "HEAD", "POST", "OPTIONS"])
def proxy():
    if request.method == "OPTIONS":
        return Response(status=204)

    target_url = request.args.get("url", type=str)

    try:
        target_url = _validate_target(target_url)
    except ValueError as exc:
        log.warning("Blocked proxy request: %s", exc)
        return jsonify({"error": str(exc)}), 400

    try:
        parsed = urlparse(target_url)
        log.info("%s %s%s", request.method, parsed.hostname, parsed.path)

        upstream = _upstream_request(target_url)
        try:
            return _build_browser_response(upstream)
        finally:
            upstream.close()

    except requests.Timeout:
        log.warning("Upstream timeout: %s", target_url)
        return jsonify(
            {
                "error": "Upstream request timed out",
                "type": "timeout",
            }
        ), 504

    except requests.TooManyRedirects as exc:
        log.warning("Redirect error: %s", exc)
        return jsonify(
            {
                "error": str(exc),
                "type": "redirect_error",
            }
        ), 502

    except requests.RequestException as exc:
        log.exception("Upstream request failed")
        return jsonify(
            {
                "error": str(exc),
                "type": "upstream_error",
            }
        ), 502

    except Exception as exc:
        log.exception("Unexpected proxy failure")
        return jsonify(
            {
                "error": str(exc),
                "type": "proxy_error",
            }
        ), 500


@app.route("/")
def serve_index():
    # Prefer the configured filename. If it does not exist, tolerate common
    # download/copy suffixes so the server is less brittle during testing.
    candidates = [
        INDEX_FILE,
        "sec-data-app.html",
        "sec-data-app(3).html",
        "index.html",
    ]

    for candidate in candidates:
        path = os.path.join(BASE_DIR, candidate)
        if os.path.isfile(path):
            response = send_from_directory(BASE_DIR, candidate)
            response.headers["Cache-Control"] = "no-store"
            return response

    return jsonify(
        {
            "error": "Dashboard HTML file not found",
            "searched": candidates,
            "directory": BASE_DIR,
        }
    ), 404


@app.route("/<path:path>")
def serve_static(path):
    response = send_from_directory(BASE_DIR, path)

    # Avoid stale JavaScript/HTML while actively developing the single-file app.
    if path.lower().endswith((".html", ".js", ".css")):
        response.headers["Cache-Control"] = "no-store"

    return response


if __name__ == "__main__":
    print()
    print("SEC Dashboard Local Proxy 2.0")
    print(f"Dashboard: http://localhost:{PORT}")
    print(f"Health:    http://localhost:{PORT}/health")
    print(f"Binding:   {HOST}:{PORT}")
    print()
    print("Allowed upstream hosts:")
    for h in sorted(ALLOWED_HOSTS):
        print(f"  - {h}")
    print()

    # threaded=True lets simultaneous dashboard API calls proceed independently.
    app.run(
        host=HOST,
        port=PORT,
        threaded=True,
        debug=False,
        use_reloader=False,
    )
