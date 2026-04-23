#!/bin/bash
# 將程式註冊到應用程式選單（只需執行一次）
set -e

BASE="$(cd "$(dirname "$0")" && pwd)"
DESKTOP_FILE="$HOME/.local/share/applications/biring-label-printer.desktop"

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

# 更新應用程式資料庫
update-desktop-database "$HOME/.local/share/applications" 2>/dev/null || true

# 同時在桌面建立捷徑（可選）
DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/桌面")"
if [ -d "$DESKTOP_DIR" ]; then
    cp "$DESKTOP_FILE" "$DESKTOP_DIR/必榮標籤列印.desktop"
    chmod +x "$DESKTOP_DIR/必榮標籤列印.desktop"
    # GNOME 需標記為可信任
    gio set "$DESKTOP_DIR/必榮標籤列印.desktop" metadata::trusted true 2>/dev/null || true
fi

echo "[完成] 已註冊到應用程式選單"
echo "       選單中搜尋「必榮」即可啟動"
[ -d "$DESKTOP_DIR" ] && echo "       桌面也已建立啟動圖示"
