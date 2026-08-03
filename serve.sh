#!/bin/bash
# Local test server for Home Dashboard.
# Usage:  ./serve.sh          -> http://localhost:8000
#         ./serve.sh 9000     -> http://localhost:9000
cd "$(dirname "$0")" || exit 1
PORT="${1:-8000}"
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null)

echo "Home Dashboard serving from: $(pwd)"
echo "  Mac browser : http://localhost:$PORT"
[ -n "$IP" ] && echo "  Tablet (LAN): http://$IP:$PORT   (no service worker over plain http)"
echo "  Ctrl+C to stop"
echo

command -v open >/dev/null && (sleep 1 && open "http://localhost:$PORT") &
python3 -m http.server "$PORT" --bind 0.0.0.0
