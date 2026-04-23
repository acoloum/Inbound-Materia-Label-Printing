#!/bin/bash
cd "$(dirname "$0")"

# 安裝必要系統套件（tkinter + venv + 中文字型）
MISSING_APT=()
python3 -c "import tkinter" 2>/dev/null || MISSING_APT+=(python3-tk)
python3 -m venv --help >/dev/null 2>&1   || MISSING_APT+=(python3-venv)
fc-list 2>/dev/null | grep -qi "NotoSansCJK" || MISSING_APT+=(fonts-noto-cjk)

if [ ${#MISSING_APT[@]} -gt 0 ]; then
    echo "安裝系統套件：${MISSING_APT[*]}"
    sudo apt install -y "${MISSING_APT[@]}"
fi

# 建立虛擬環境（不加 --system-site-packages，避免系統舊版 PIL 干擾）
VENV_DIR="$(dirname "$0")/.venv"
if [ ! -d "$VENV_DIR/bin/pip" ]; then
    echo "建立虛擬環境..."
    python3 -m venv "$VENV_DIR" || {
        echo "[錯誤] 虛擬環境建立失敗"
        exit 1
    }
fi

# 在虛擬環境內安裝 Python 套件（含完整 Pillow 與 ImageTk）
"$VENV_DIR/bin/pip" install --quiet pycups "qrcode[pil]" pillow openpyxl

# 啟動程式
"$VENV_DIR/bin/python3" run.py
