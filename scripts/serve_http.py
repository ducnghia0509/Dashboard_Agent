# -*- coding: utf-8 -*-
"""Chay 1 MCP server qua HTTP/SSE (de OpenClaw gateway trong container goi qua
host.docker.internal). Dung cho che do gateway; che do Claude Code local van dung
stdio qua mcp_config.json.

    python scripts/serve_http.py ingest 8810
    python scripts/serve_http.py qa 8811
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(_ROOT, ".env"))


def main():
    if len(sys.argv) < 3:
        print("usage: serve_http.py <ingest|qa> <port>", file=sys.stderr)
        sys.exit(2)
    which, port = sys.argv[1], int(sys.argv[2])
    if which == "ingest":
        from servers.ingest_server import mcp
    elif which == "qa":
        from servers.qa_server import mcp
    else:
        print(f"unknown server '{which}'", file=sys.stderr)
        sys.exit(2)

    mcp.settings.host = os.environ.get("MCP_BIND_HOST", "0.0.0.0")
    mcp.settings.port = port

    # OpenClaw gateway chay trong container, goi vao day qua host.docker.internal ->
    # Host header != localhost nen bi chan boi DNS-rebinding protection (HTTP 421).
    # Cho phep them cac host tin cay (host.docker.internal + docker bridge gateway).
    from mcp.server.transport_security import TransportSecuritySettings

    extra = os.environ.get("MCP_ALLOWED_HOSTS", "host.docker.internal:*,172.17.0.1:*,172.18.0.1:*")
    extra_hosts = [h.strip() for h in extra.split(",") if h.strip()]
    mcp.settings.transport_security = TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=["127.0.0.1:*", "localhost:*", "[::1]:*", *extra_hosts],
        allowed_origins=[
            "http://127.0.0.1:*", "http://localhost:*", "http://[::1]:*",
            *[f"http://{h}" for h in extra_hosts],
        ],
    )
    print(f"[serve_http] {which} on {mcp.settings.host}:{port} (sse), allowed_hosts={extra_hosts}", flush=True)
    mcp.run(transport="sse")


if __name__ == "__main__":
    main()
