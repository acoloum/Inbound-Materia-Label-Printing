#!/bin/bash
# 將程式註冊到應用程式選單與桌面
set -e

BASE="$(cd "$(dirname "$0")" && pwd)"
APP_ID="biring-label-printer"
DESKTOP_FILE="$HOME/.local/share/applications/${APP_ID}.desktop"

mkdir -p "$(dirname "$DESKTOP_FILE")"

cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=必榮進料標籤列印
Name[zh_TW]=必榮進料標籤列印
Comment=進料標籤列印系統
Exec=bash "$BASE/start.sh"
Path=$BASE
Icon=printer
Terminal=false
Categories=Office;Utility;
StartupNotify=true
EOF

chmod +x "$DESKTOP_FILE"
chmod +x "$BASE/start.sh"

# 刷新應用程式資料庫
update-desktop-database "$HOME/.local/share/applications" 2>/dev/null || true

# 同時在桌面建立捷徑
DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/桌面")"
if [ -d "$DESKTOP_DIR" ]; then
    DESKTOP_COPY="$DESKTOP_DIR/必榮標籤列印.desktop"
    cp "$DESKTOP_FILE" "$DESKTOP_COPY"
    chmod +x "$DESKTOP_COPY"
    # GNOME 43+：標記為可信任（否則雙擊無反應）
    gio set "$DESKTOP_COPY" metadata::trusted true 2>/dev/null || true
    # 舊版 Nautilus 備用
    gio set -t string "$DESKTOP_COPY" metadata::nautilus-trusted true 2>/dev/null || true
fi

echo "[完成] 已註冊到應用程式選單"
echo "       選單中搜尋「必榮」即可啟動"
if [ -d "$DESKTOP_DIR" ]; then
    echo "       桌面圖示已建立"
    echo ""
    echo "★ 若桌面圖示雙擊無反應："
    echo "  1. 右鍵點圖示 → 選「允許啟動」(Allow Launching)"
    echo "  2. 或直接從左下角選單搜尋「必榮」啟動"
fi
