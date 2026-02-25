#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="${ROOT_DIR}/build"
XLSX_FILE="${OUT_DIR}/swale_calculator.xlsx"
VENV_DIR="${ROOT_DIR}/.venv-linux"

cd "${ROOT_DIR}"
mkdir -p "${OUT_DIR}"

# Create venv if missing
if [[ ! -x "${VENV_DIR}/bin/python" ]]; then
  python3 -m venv "${VENV_DIR}"
fi

PY="${VENV_DIR}/bin/python"

# Install deps if needed (simple sentinel)
if [[ ! -f "${VENV_DIR}/.deps_installed" ]]; then
  "${PY}" -m pip install --upgrade pip >/dev/null
  "${PY}" -m pip install -r "${ROOT_DIR}/requirements.txt"
  touch "${VENV_DIR}/.deps_installed"
fi

# Generate workbook
"${PY}" swale-calculator.py --out "${XLSX_FILE}"
