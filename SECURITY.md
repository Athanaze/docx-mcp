# Security

## Reporting a vulnerability

Please report security issues privately (open a GitHub Security Advisory or email the maintainers) instead of using public issues.

## Deployment risks

This MCP server can **read and write Word documents** on the machine (or container) where it runs.

- **Local use (typical)**: If the MCP server runs on the **same machine** as the client (stdio with Cursor/Claude Desktop, or HTTP on `127.0.0.1`), **transport encryption is not this project’s job** and is usually unnecessary—the process boundary is your OS and host app. No TLS or “end-to-end encryption” layer is provided or required for that setup.
- **`DOCUMENT_ROOT`**: When set, all document paths are resolved inside this directory. Use it in Docker and shared hosting so clients cannot escape the intended folder.
- **HTTP on untrusted networks / the public internet**: Streamable HTTP (`MCP_TRANSPORT=streamable-http` or the ASGI app) has **no built-in authentication** and speaks plain HTTP unless **you** terminate TLS in front of it (reverse proxy, mesh, etc.). Only expose it beyond localhost if you add that perimeter yourself; this server does not implement application-level crypto for MCP traffic.
- **Malicious documents**: Parsing and editing untrusted `.docx` files has the same risks as any document tooling (embedded objects, resource exhaustion). Prefer size limits and trusted sources in automated workflows.
- **Secrets in env**: Avoid putting passwords for `encrypt_document` / `decrypt_document` in shared logs or process listings.

## Supply chain

Install from PyPI or a pinned git tag. Verify checksums and lockfiles in production images.
