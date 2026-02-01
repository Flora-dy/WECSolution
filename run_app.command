#!/bin/zsh
# Double-clickable launcher for the Streamlit app with an isolated venv
SCRIPT_DIR="$(cd -- "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

VENV_DIR="$SCRIPT_DIR/.venv"

# Create venv if missing
if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
fi

export PYTHONNOUSERSITE=1
export PYTHONPATH=""
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements-app.txt

# Launch from the venv to avoid system-site conflicts
exec "$VENV_DIR/bin/python" -m streamlit run app.py
