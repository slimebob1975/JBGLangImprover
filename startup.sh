#!/bin/bash

# Installera Linux-paket här om det saknas
echo "[INFO] No known missing Linux packages..."

# Starta appen
echo "[INFO] Starting JBGLangImrprover API via Gunicorn"

# Kör med 4 workers och uvicorn workers
exec gunicorn app.main:app \
    --workers 2 \
    --worker-class uvicorn.workers.UvicornWorker \
    --bind 0.0.0.0:8000 \
    --timeout 600
