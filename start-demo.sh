#!/usr/bin/env bash
set -euo pipefail
cd -- "$(dirname -- "${BASH_SOURCE[0]}")"
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
export QO_DEMO=1 QO_ENGINE=demo QO_DATA_DIR=data/demo QO_PUBLIC_ORIGIN=http://127.0.0.1:8765
unset QO_DATABASE_URL QO_USERS_FILE QO_CLAMD QO_OIDC_ISSUER QO_OIDC_CLIENT_ID QO_OIDC_CLIENT_SECRET
exec .venv/bin/python -m uvicorn server.app:app --host 127.0.0.1 --port 8765 --no-access-log
