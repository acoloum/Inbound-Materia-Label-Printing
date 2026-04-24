#!/bin/bash
# 首次設定腳本 — 需從終端機執行一次
set -e
cd "$(dirname "$0")"

echo "=== 必榮標籤列印 — 首次設定 ==="

# 安裝必要系統套件
MISSING_APT=()
python3 -c "import tkinter" 2>/dev/null || MISSING_APT+=(python3-tk)
python3 -m venv --help >/dev/null 2>&1   || MISSING_APT+=(python3-venv)
fc-list 2>/dev/null | grep -qi "NotoSansCJK" || MISSING_APT+=(fonts-noto-cjk)
# reportlab 無法使用 Noto（CFF outlines），需 WQY MicroHei（TrueType outlines）
fc-list 2>/dev/null | grep -qi "WenQuanYi Micro Hei" || MISSING_APT+=(fonts-wqy-microhei)
dpkg -s libcups2-dev >/dev/null 2>&1 || MISSING_APT+=(libcups2-dev)
dpkg -s python3-dev  >/dev/null 2>&1 || MISSING_APT+=(python3-dev)
command -v gcc >/dev/null 2>&1        || MISSING_APT+=(gcc)

if [ ${#MISSING_APT[@]} -gt 0 ]; then
    echo "→ 安裝系統套件：${MISSING_APT[*]}"
    sudo apt install -y "${MISSING_APT[@]}"
fi

# 建立虛擬環境
VENV_DIR="$(pwd)/.venv"
if [ ! -f "$VENV_DIR/bin/pip" ]; then
    echo "→ 建立虛擬環境..."
    rm -rf "$VENV_DIR"
    python3 -m venv "$VENV_DIR"
fi

# 安裝 Python 套件
echo "→ 安裝 Python 套件..."
"$VENV_DIR/bin/pip" install --quiet --upgrade pip
"$VENV_DIR/bin/pip" install --quiet pycups "qrcode[pil]" pillow openpyxl reportlab

# 註冊桌面啟動檔
echo "→ 註冊桌面啟動檔..."
bash "$(pwd)/install-desktop.sh"

echo ""
echo "=== 完成 ==="
echo "之後從應用程式選單搜尋「必榮」即可啟動"
