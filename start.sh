#!/bin/bash
cd "$(dirname "$0")"

# tkinter 必須透過 apt 安裝（無法放進 venv）
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "安裝 tkinter..."
    sudo apt install -y python3-tk
fi

# 建立虛擬環境（--system-site-packages 讓 venv 可存取 tkinter）
VENV_DIR="$(dirname "$0")/.venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "建立虛擬環境..."
    python3 -m venv "$VENV_DIR" --system-site-packages
fi

# 在虛擬環境內安裝套件
"$VENV_DIR/bin/pip" install --quiet pycups "qrcode[pil]" pillow openpyxl

# 確認 Noto CJK 字型
fc-list 2>/dev/null | grep -qi "NotoSansCJK" || \
    echo "建議執行：sudo apt install fonts-noto-cjk"

# 啟動程式
"$VENV_DIR/bin/python3" run.py
