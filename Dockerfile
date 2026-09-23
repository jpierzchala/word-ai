FROM python:3.12-slim@sha256:78387bc3881b8273120a12ebe6c1ab22b018ccc2c9adf565ae1ac9b536e184ea

LABEL org.opencontainers.image.title="Word AI secure live"
LABEL org.opencontainers.image.description="Live Word editing with revocable document-session consent"
LABEL org.opencontainers.image.source="https://github.com/jpierzchala/word-ai"
LABEL org.opencontainers.image.licenses="AGPL-3.0-or-later"

ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# Only the audited profile enters this image; no legacy tools or host documents.
COPY word_ai_mcp/secure_live.py /app/secure_live.py
COPY word_ai_mcp/document_state.py /app/document_state.py
COPY word_ai_mcp/image_assets.py /app/image_assets.py
COPY secure-addin/taskpane.html secure-addin/taskpane.js secure-addin/safety.js secure-addin/document-edit.js secure-addin/live-ops.js secure-addin/section-move.js secure-addin/taskpane.css /app/static/
COPY office-addin/assets/icon-32.png office-addin/assets/icon-64.png /app/static/
COPY LICENSE /app/LICENSE
USER 10001:10001
EXPOSE 3100
ENTRYPOINT ["python", "/app/secure_live.py"]
CMD ["serve"]
