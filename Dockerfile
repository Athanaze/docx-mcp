# Word Document MCP Server — Streamable HTTP via ASGI (Uvicorn)
FROM python:3.12-slim-bookworm

RUN useradd --create-home --uid 1000 --shell /bin/bash mcp

WORKDIR /app

COPY pyproject.toml README.md LICENSE ./
COPY word_document_server ./word_document_server/
COPY word_mcp_server.py ./

RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir ".[asgi]"

USER mcp

ENV DOCUMENT_ROOT=/workspace
ENV MCP_PATH=/mcp

EXPOSE 8000
VOLUME ["/workspace"]

# Production-style HTTP MCP (Streamable HTTP). Mount a volume at /workspace for .docx files.
CMD ["uvicorn", "word_document_server.asgi:app", "--host", "0.0.0.0", "--port", "8000"]
