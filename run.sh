#!/usr/bin/env bash
# Launch the CMC Report Automation desktop app from source on Linux.
# Usage:  ./run.sh
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

if [ -f ".venv/bin/python" ]; then
    PY=".venv/bin/python"
else
    PY="python3"
fi

exec "$PY" -m app.main "$@"
