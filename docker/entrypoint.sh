#!/bin/sh
set -eu

display="${DISPLAY:-:99}"
screen_width="${SCREEN_WIDTH:-1280}"
screen_height="${SCREEN_HEIGHT:-720}"
screen_depth="${SCREEN_DEPTH:-24}"

Xvfb "$display" -screen 0 "${screen_width}x${screen_height}x${screen_depth}" -nolisten tcp &
xvfb_pid=$!

cleanup() {
    kill "$xvfb_pid" 2>/dev/null || true
}
trap cleanup EXIT INT TERM

attempt=0
until xdpyinfo -display "$display" >/dev/null 2>&1; do
    attempt=$((attempt + 1))
    if [ "$attempt" -ge 50 ]; then
        echo "Xvfb did not become ready." >&2
        exit 1
    fi
    sleep 0.1
done

x11vnc -display "$display" -forever -shared -nopw -rfbport 5900 \
    -o /tmp/x11vnc.log -bg
websockify --web=/usr/share/novnc/ 6080 localhost:5900 &

echo "CoinTex GUI is available through noVNC on port 6080."
exec "$@"
