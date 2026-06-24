#!/usr/bin/env bash
# Install dependencies and run the interactive setup (Linux/macOS).
set -e
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
python3 qobuz_rpc_cli.py --setup
