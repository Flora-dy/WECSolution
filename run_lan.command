#!/bin/zsh
set -euo pipefail

cd "$(dirname "$0")"

if [[ -x ".venv/bin/python" ]]; then
  PY=".venv/bin/python"
else
  PY="python3"
fi

# 使用非特权端口（>=1024），避免 macOS 报 PermissionError
exec "$PY" -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
