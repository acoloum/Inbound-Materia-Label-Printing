#!/bin/bash
# 純啟動腳本 — 供桌面圖示點擊使用（不需 sudo）
cd "$(dirname "$0")"

LOG="$HOME/.biring-label-printer.log"
exec > >(tee -a "$LOG") 2>&1
echo ""
echo "=== 啟動 $(date '+%Y-%m-%d %H:%M:%S') ==="

VENV_DIR="$(pwd)/.venv"

# 若未完成首次設定，顯示提示視窗
if [ ! -f "$VENV_DIR/bin/python3" ]; then
    MSG="請先從終端機執行：./setup.sh 完成首次設定"
    zenity --error --title="必榮標籤列印" --text="$MSG" 2>/dev/null \
        || notify-send "必榮標籤列印" "$MSG" 2>/dev/null \
        || echo "$MSG"
    exit 1
fi

# 啟動程式
exec "$VENV_DIR/bin/python3" run.py
