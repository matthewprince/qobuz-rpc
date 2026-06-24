#!/usr/bin/env bash
# Build standalone Linux binaries with PyInstaller.
# Note: PyInstaller binaries are tied to the glibc of the build machine; for
# wide distribution, `pipx install` from source is usually friendlier.
set -e
cd "$(dirname "$0")"

echo "[*] Installing PyInstaller..."
python3 -m pip install pyinstaller

echo "[*] Building CLI (dist/QobuzRPC-CLI)..."
python3 -m PyInstaller --noconfirm --onefile --console \
    --name "QobuzRPC-CLI" \
    --collect-all "jeepney" \
    qobuz_rpc_cli.py

echo "[*] Building GUI (dist/QobuzRPC)..."
python3 -m PyInstaller --noconfirm --onefile --windowed \
    --name "QobuzRPC" \
    --add-data "icon.png:." \
    --add-data "config.example.json:." \
    --collect-all "jeepney" \
    qobuz_rpc.py

echo
echo "[+] Done. Binaries in dist/. Keep icon.png and config.example.json alongside them."
