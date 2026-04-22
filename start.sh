#!/bin/bash
cd "$(dirname "$0")"

python3 -c "import cups, PIL, qrcode, openpyxl" 2>/dev/null || {
    echo "安裝必要套件..."
    pip3 install pycups "qrcode[pil]" pillow openpyxl
}

# 確認 Noto CJK 字型已安裝
fc-list | grep -qi "NotoSansCJK" 2>/dev/null || {
    echo "建議安裝中文字型：sudo apt install fonts-noto-cjk"
}

python3 run.py
