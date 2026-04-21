#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
必榮 進料標籤列印系統
跨電腦通用版本 - 不需要授權碼，任何 Windows 電腦皆可執行
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import sys
import io
import win32print
import win32ui
import win32con
import win32gui
from PIL import Image, ImageDraw, ImageFont
import qrcode
import openpyxl
from datetime import datetime

# ── 設定 ──────────────────────────────────────────────────────────────────────
APP_TITLE = "必榮 進料標籤列印系統"
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "FastReport_sqllite.db")

# 標籤尺寸 (mm) — 實際標籤紙尺寸
LABEL_W_MM = 99.86
LABEL_H_MM = 59.3

# 四周留白 (mm)
MARGIN_MM = 2.5

# 列印解析度 DPI
PRINT_DPI = 203

# 標籤像素尺寸（含留白）
LABEL_W_PX = int(LABEL_W_MM / 25.4 * PRINT_DPI)
LABEL_H_PX = int(LABEL_H_MM / 25.4 * PRINT_DPI)

# 留白像素
MARGIN_PX = int(MARGIN_MM / 25.4 * PRINT_DPI)

# 字型路徑（微軟正黑體）
FONT_PATH = "C:/Windows/Fonts/msjh.ttc"
FONT_BOLD_PATH = "C:/Windows/Fonts/msjhbd.ttc"

# 表格欄位定義
COLUMNS = [
    ("序號", 70), ("供應商名稱", 80), ("訂單編號", 90),
    ("材質", 50), ("尺寸", 50), ("批號", 60), ("特殊", 60),
    ("長度", 60), ("數量", 50), ("製造編號/爐號", 120), ("進貨日期", 90),
]


# ── 資料庫 ────────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS MYTABLE (
            INDX INT, SN INT, QTY INT, SEL NUM,
            序號 TEXT, 供應商名稱 TEXT, 訂單編號 TEXT, 材質 TEXT,
            尺寸 TEXT, 批號 TEXT, 特殊 TEXT, 長度 TEXT,
            數量 INT, 重量 TEXT, 不良支數 TEXT, 樣品量 TEXT,
            檢驗尺寸 TEXT, 檢驗外觀 TEXT, 檢驗材質 TEXT, 判定 TEXT,
            製造編號爐號 TEXT, 進貨日期 TEXT, 備註 TEXT,
            PKGQTY INT, SNN INT
        )
    """)
    conn.commit()
    conn.close()


def load_table_data():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT 序號, 供應商名稱, 訂單編號, 材質, 尺寸, 批號, 特殊,
                   長度, 數量, "製造編號/爐號", 進貨日期, PKGQTY, SNN, SN
            FROM MYTABLE ORDER BY SN
        """)
        return cur.fetchall()
    except Exception:
        return []
    finally:
        conn.close()


def import_excel_to_db(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]

    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM MYTABLE")

    sn = 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue

        def get(col_name):
            try:
                idx = headers.index(col_name)
                v = row[idx]
                if v is None:
                    return ""
                if isinstance(v, datetime):
                    return v.strftime("%Y/%m/%d")
                return str(v)
            except (ValueError, IndexError):
                return ""

        cur.execute("""
            INSERT INTO MYTABLE
            (SN, QTY, SEL, 序號, 供應商名稱, 訂單編號, 材質, 尺寸, 批號, 特殊,
             長度, 數量, 重量, 不良支數, 樣品量, 檢驗尺寸, 檢驗外觀, 檢驗材質,
             判定, "製造編號/爐號", 進貨日期, 備註, PKGQTY, SNN)
            VALUES (?,1,1,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,NULL)
        """, (
            sn,
            get("序號"), get("供應商名稱"), get("訂單編號"), get("材質"),
            get("尺寸"), get("批號"), get("特殊"), get("長度"), get("數量"),
            get("重量"), get("不良支數"), get("樣品量"), get("檢驗尺寸"),
            get("檢驗外觀"), get("檢驗材質"), get("判定"), get("製造編號/爐號"),
            get("進貨日期"), get("備註"),
        ))
        sn += 1

    conn.commit()
    conn.close()
    return sn - 1


# ── 標籤圖像產生 ───────────────────────────────────────────────────────────────

def _load_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def _draw_cell(draw, x1, y1, x2, y2, text, font, h_align="center", pad=6):
    """在格子內置中（或靠左）繪製文字，自動截斷過長文字"""
    cell_w = x2 - x1
    cell_h = y2 - y1
    # 若文字過長則逐字截斷
    while text:
        bbox = font.getbbox(text)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        if tw <= cell_w - pad * 2:
            break
        text = text[:-1]

    bbox = font.getbbox(text)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]

    if h_align == "center":
        tx = x1 + (cell_w - tw) // 2
    else:
        tx = x1 + pad

    # bbox[1] 是字體頂部相對於繪製原點的偏移（通常是負值），
    # 需減去它才能讓文字視覺上垂直置中
    ty = y1 + (cell_h - th) // 2 - bbox[1]
    draw.text((tx, ty), text, fill="black", font=font)


def make_label_image(record, pkg_no=1, pkg_total=1):
    """產生一張標籤的 PIL Image，四周保留 MARGIN_PX 留白"""
    W, H = LABEL_W_PX, LABEL_H_PX
    M = MARGIN_PX  # 四周留白像素

    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    font_lbl  = _load_font(FONT_BOLD_PATH, 36)
    font_data = _load_font(FONT_BOLD_PATH, 36)
    font_bot  = _load_font(FONT_BOLD_PATH, 32)  # 底部略小，確保訂單編號不截斷

    # ── 可用內容區域（扣除四周留白）─────────────────────────────────────────
    CX = M          # 內容起始 X
    CY = M          # 內容起始 Y
    CW = W - 2 * M  # 內容寬度
    CH = H - 2 * M  # 內容高度

    # 8 資料行 + 1 底部行
    BOT_H  = int(CH * 0.09)
    MAIN_H = CH - BOT_H
    ROW_H  = MAIN_H // 8
    MAIN_H = ROW_H * 8

    # 欄寬比例：標籤欄 33% / QR欄 26% / 資料欄 = 剩餘
    LBL_W = int(CW * 0.331)
    QR_W  = int(CW * 0.260)
    X_LBL  = CX
    X_DATA = CX + LBL_W
    X_QR   = CX + CW - QR_W
    X_END  = CX + CW
    QR_ROWS = 4

    # ── QR Code ──────────────────────────────────────────────────────────────
    qr_text = (
        f"供應商 : {record.get('供應商名稱','')}\n"
        f"進貨日期 : {record.get('進貨日期','')}\n"
        f"材質/特殊 : {record.get('材質','')}{record.get('特殊','')}\n"
        f"尺寸 : {record.get('尺寸','')}\n"
        f"長度 : {record.get('長度','')}\n"
        f"批號 : {record.get('批號','')}\n"
        f"數量 : {pkg_no}/{pkg_total}\n"
        f"製造編號/爐號 : {record.get('製造編號/爐號','')}\n"
        f"ERP序號 : {record.get('序號','')}\n"
        f"訂單編號 : {record.get('訂單編號','')}"
    )
    qr = qrcode.QRCode(version=None,
                       error_correction=qrcode.constants.ERROR_CORRECT_L,
                       box_size=3, border=1)
    qr.add_data(qr_text)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_size = min(QR_W - 4, ROW_H * QR_ROWS - 4)
    qr_img = qr_img.resize((qr_size, qr_size), Image.LANCZOS)
    qr_px = X_QR + (QR_W - qr_size) // 2
    qr_py = CY + (ROW_H * QR_ROWS - qr_size) // 2
    img.paste(qr_img, (qr_px, qr_py))

    # ── 8 行資料 ──────────────────────────────────────────────────────────────
    rows_def = [
        ("供應商",        str(record.get("供應商名稱") or "")),
        ("進貨日期",      str(record.get("進貨日期") or "")),
        ("材質/特殊",     f"{record.get('材質','')}{record.get('特殊','')}"),
        ("尺寸",          str(record.get("尺寸") or "")),
        ("長度",          str(record.get("長度") or "")),
        ("批號",          str(record.get("批號") or "")),
        ("數量",          f"{pkg_no}/{pkg_total}"),
        ("製造編號/爐號", str(record.get("製造編號/爐號") or "")),
    ]

    for i, (lbl, val) in enumerate(rows_def):
        y0 = CY + i * ROW_H
        y1 = y0 + ROW_H
        x_right = X_QR if i < QR_ROWS else X_END

        if i > 0:
            line_x = X_END if i >= QR_ROWS else x_right
            draw.line([(CX, y0), (line_x, y0)], fill="black", width=1)

        draw.line([(X_DATA, y0), (X_DATA, y1)], fill="black", width=1)
        _draw_cell(draw, X_LBL, y0, X_DATA, y1, lbl, font_lbl, "center")
        _draw_cell(draw, X_DATA, y0, x_right, y1, val, font_data, "left")

    draw.line([(X_QR, CY), (X_QR, CY + ROW_H * QR_ROWS)], fill="black", width=1)

    # ── 底部行 ────────────────────────────────────────────────────────────────
    BY = CY + MAIN_H
    BH = BOT_H
    draw.line([(CX, BY),      (X_END, BY)],      fill="black", width=2)
    draw.line([(CX, BY + BH), (X_END, BY + BH)], fill="black", width=2)

    # 底部欄：標籤欄 22%、資料欄 28%（各兩組，共 100%）
    b1 = CX + int(CW * 0.22)   # ERP序號 標籤右緣
    b2 = CX + int(CW * 0.50)   # ERP序號 值右緣（同時是訂單編號標籤左緣）
    b3 = CX + int(CW * 0.72)   # 訂單編號 標籤右緣
    bx = [CX, b1, b2, b3, X_END]
    for x in bx[1:-1]:
        draw.line([(x, BY), (x, BY + BH)], fill="black", width=1)

    _draw_cell(draw, bx[0], BY, bx[1], BY + BH, "ERP序號",                           font_bot, "center")
    _draw_cell(draw, bx[1], BY, bx[2], BY + BH, str(record.get("序號") or ""),       font_bot, "center")
    _draw_cell(draw, bx[2], BY, bx[3], BY + BH, "訂單編號",                          font_bot, "center")
    _draw_cell(draw, bx[3], BY, bx[4], BY + BH, str(record.get("訂單編號") or ""),   font_bot, "center")

    # ── 外框（在留白內緣畫）────────────────────────────────────────────────────
    draw.rectangle([(CX, CY), (X_END, BY + BH)], outline="black", width=2)

    return img


# ── Windows 列印 ──────────────────────────────────────────────────────────────

def print_label(printer_name, pil_image):
    """透過 Windows GDI 將 PIL Image 送到印表機"""
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)

        pw = hdc.GetDeviceCaps(win32con.HORZRES)
        ph = hdc.GetDeviceCaps(win32con.VERTRES)

        # 縮放到印表機解析度
        img_resized = pil_image.resize((pw, ph), Image.LANCZOS)

        dib = img_resized.convert("RGB")

        hdc.StartDoc("進料標籤")
        hdc.StartPage()

        import win32ui as wui
        dib_bmp = wui.CreateBitmap()
        dib_bmp.CreateCompatibleBitmap(hdc, pw, ph)

        mem_dc = hdc.CreateCompatibleDC()
        mem_dc.SelectObject(dib_bmp)

        bmp_info = dib.tobytes("raw", "BGRX")
        dib_bmp.SetBitmapBits(bmp_info)

        hdc.BitBlt((0, 0), (pw, ph), mem_dc, (0, 0), win32con.SRCCOPY)

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()
        mem_dc.DeleteDC()
    finally:
        win32print.ClosePrinter(hprinter)


def print_label_simple(printer_name, pil_image):
    """強制以 LABEL_W_MM×LABEL_H_MM 為單張紙尺寸列印，避免驅動把兩張標籤當成一頁。"""
    # DEVMODE 欄位旗標
    DM_ORIENTATION = 0x00000001
    DM_PAPERSIZE   = 0x00000002
    DM_PAPERLENGTH = 0x00000004
    DM_PAPERWIDTH  = 0x00000008
    DMPAPER_USER   = 256  # 自訂紙張

    hprinter = win32print.OpenPrinter(printer_name)
    try:
        props = win32print.GetPrinter(hprinter, 2)
        devmode = props["pDevMode"]
        # 單位：0.1mm
        devmode.PaperSize   = DMPAPER_USER
        devmode.PaperWidth  = int(round(LABEL_W_MM * 10))
        devmode.PaperLength = int(round(LABEL_H_MM * 10))
        devmode.Orientation = 1  # Portrait
        devmode.Fields = (devmode.Fields
                          | DM_PAPERSIZE | DM_PAPERWIDTH
                          | DM_PAPERLENGTH | DM_ORIENTATION)

        # 以這份 DEVMODE 建立 DC
        hdc_int = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
        hdc = win32ui.CreateDCFromHandle(hdc_int)
    finally:
        win32print.ClosePrinter(hprinter)

    # 取得此 DC 的實際可列印像素數
    dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
    dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
    page_w = hdc.GetDeviceCaps(win32con.HORZRES)
    page_h = hdc.GetDeviceCaps(win32con.VERTRES)

    # 依標籤實際 mm 換算的像素
    px_w = int(LABEL_W_MM / 25.4 * dpi_x)
    px_h = int(LABEL_H_MM / 25.4 * dpi_y)

    # 若驅動回報的可列印區域小於標籤（通常因不可列印邊界），縮到剛好填滿
    if page_w and page_h:
        px_w = min(px_w, page_w)
        px_h = min(px_h, page_h)

    img_resized = pil_image.resize((px_w, px_h), Image.LANCZOS).convert("RGB")

    hdc.StartDoc("進料標籤")
    hdc.StartPage()

    from PIL.ImageWin import Dib
    dib = Dib(img_resized)
    dib.draw(hdc.GetHandleOutput(), (0, 0, px_w, px_h))

    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


# ── 主視窗 ────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x700")
        self.resizable(True, True)
        self._selected_ids = set()
        self._build_ui()
        self._refresh_table()

    # ── UI 建構 ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── 工具列 ────────────────────────────────────────────────────────────
        toolbar = ttk.Frame(self, padding=4)
        toolbar.pack(fill="x", side="top")

        ttk.Button(toolbar, text="📂 匯入 Excel", command=self._import_excel).pack(side="left", padx=2)
        ttk.Separator(toolbar, orient="vertical").pack(side="left", padx=6, fill="y")

        ttk.Label(toolbar, text="印表機:").pack(side="left")
        self._printer_var = tk.StringVar()
        self._printer_cb = ttk.Combobox(toolbar, textvariable=self._printer_var, width=30, state="readonly")
        self._printer_cb.pack(side="left", padx=4)
        self._refresh_printers()
        ttk.Button(toolbar, text="🔄", command=self._refresh_printers, width=3).pack(side="left")

        ttk.Separator(toolbar, orient="vertical").pack(side="left", padx=6, fill="y")
        ttk.Button(toolbar, text="☑ 全選", command=self._select_all).pack(side="left", padx=2)
        ttk.Button(toolbar, text="☐ 取消全選", command=self._deselect_all).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", padx=6, fill="y")
        ttk.Button(toolbar, text="🔍 預覽標籤", command=self._preview_label).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🖨 列印選取", command=self._print_selected,
                   style="Accent.TButton").pack(side="left", padx=2)
        ttk.Button(toolbar, text="🔢 指定張數", command=self._print_range).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", padx=6, fill="y")
        ttk.Button(toolbar, text="🗑 清除資料", command=self._clear_data).pack(side="left", padx=2)

        # ── 資料表格 ──────────────────────────────────────────────────────────
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=6, pady=4)

        cols = ("chk",) + tuple(c[0] for c in COLUMNS) + ("每包數量",)
        self._tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")

        self._tree.heading("chk", text="✓")
        self._tree.column("chk", width=30, anchor="center", stretch=False)
        for name, w in COLUMNS:
            self._tree.heading(name, text=name)
            self._tree.column(name, width=w, anchor="center")
        self._tree.heading("每包數量", text="每包數量")
        self._tree.column("每包數量", width=70, anchor="center")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self._tree.bind("<ButtonRelease-1>", self._on_tree_click)

        # ── 狀態列 ────────────────────────────────────────────────────────────
        self._status = tk.StringVar(value="就緒")
        ttk.Label(self, textvariable=self._status, relief="sunken", anchor="w").pack(
            fill="x", side="bottom", padx=4, pady=2)

    # ── 輔助方法 ──────────────────────────────────────────────────────────────

    def _refresh_printers(self):
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        self._printer_cb["values"] = printers
        default = win32print.GetDefaultPrinter()
        if default in printers:
            self._printer_var.set(default)
        elif printers:
            self._printer_var.set(printers[0])

    def _refresh_table(self):
        self._tree.delete(*self._tree.get_children())
        rows = load_table_data()
        for r in rows:
            # 欄位對應: 序號,供應商名稱,訂單編號,材質,尺寸,批號,特殊,長度,數量,製造編號/爐號,進貨日期,PKGQTY,SNN,SN
            sn = r["SN"] if "SN" in r.keys() else ""
            pkg = r["PKGQTY"] if r["PKGQTY"] else 1
            vals = (
                "☐",
                r["序號"] or "", r["供應商名稱"] or "", r["訂單編號"] or "",
                r["材質"] or "", r["尺寸"] or "", r["批號"] or "", r["特殊"] or "",
                r["長度"] or "", r["數量"] or "", r["製造編號/爐號"] or "", r["進貨日期"] or "",
                pkg,
            )
            self._tree.insert("", "end", iid=str(sn), values=vals)

        cnt = len(rows)
        self._status.set(f"共 {cnt} 筆資料")

    def _on_tree_click(self, event):
        region = self._tree.identify("region", event.x, event.y)
        col = self._tree.identify_column(event.x)
        if region == "cell" and col == "#1":
            item = self._tree.identify_row(event.y)
            if item:
                self._toggle_check(item)

    def _toggle_check(self, item):
        vals = list(self._tree.item(item, "values"))
        if vals[0] == "☐":
            vals[0] = "☑"
            self._selected_ids.add(item)
        else:
            vals[0] = "☐"
            self._selected_ids.discard(item)
        self._tree.item(item, values=vals)
        cnt = len(self._selected_ids)
        self._status.set(f"已選取 {cnt} 筆")

    def _select_all(self):
        self._selected_ids.clear()
        for item in self._tree.get_children():
            vals = list(self._tree.item(item, "values"))
            vals[0] = "☑"
            self._tree.item(item, values=vals)
            self._selected_ids.add(item)
        self._status.set(f"已全選 {len(self._selected_ids)} 筆")

    def _deselect_all(self):
        for item in self._tree.get_children():
            vals = list(self._tree.item(item, "values"))
            vals[0] = "☐"
            self._tree.item(item, values=vals)
        self._selected_ids.clear()
        self._status.set("已取消全選")

    # ── 功能方法 ──────────────────────────────────────────────────────────────

    def _import_excel(self):
        path = filedialog.askopenfilename(
            title="選擇 ERP Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")]
        )
        if not path:
            return
        try:
            cnt = import_excel_to_db(path)
            self._refresh_table()
            messagebox.showinfo("匯入成功", f"成功匯入 {cnt} 筆資料")
        except Exception as e:
            messagebox.showerror("匯入失敗", str(e))

    def _get_record_by_sn(self, sn):
        conn = get_db()
        cur = conn.cursor()
        cur.execute('SELECT * FROM MYTABLE WHERE SN=?', (sn,))
        row = cur.fetchone()
        conn.close()
        return row

    def _build_record_dict(self, row):
        """將資料庫 Row 轉為 dict，相容欄位名稱差異"""
        d = dict(row)
        # 處理可能的欄位名稱差異（製造編號/爐號 vs 製造編號爐號）
        if "製造編號/爐號" not in d:
            d["製造編號/爐號"] = d.get("製造編號爐號", "")
        return d

    def _preview_label(self):
        if not self._selected_ids:
            messagebox.showwarning("提示", "請先勾選要預覽的資料")
            return
        sn = next(iter(self._selected_ids))
        row = self._get_record_by_sn(sn)
        if not row:
            messagebox.showerror("錯誤", "找不到資料")
            return
        rec = self._build_record_dict(row)
        total_qty = int(rec.get("數量") or 1)
        img = make_label_image(rec, 1, total_qty)

        # 顯示預覽視窗
        PreviewWindow(self, img)

    def _print_selected(self):
        if not self._selected_ids:
            messagebox.showwarning("提示", "請先勾選要列印的資料")
            return
        printer = self._printer_var.get()
        if not printer:
            messagebox.showwarning("提示", "請選擇印表機")
            return

        total = len(self._selected_ids)
        confirm = messagebox.askyesno("確認列印",
            f"即將列印 {total} 筆資料的標籤\n印表機：{printer}\n\n確定列印？")
        if not confirm:
            return

        success = 0
        for sn in self._selected_ids:
            try:
                row = self._get_record_by_sn(sn)
                if not row:
                    continue
                rec = self._build_record_dict(row)
                total_qty = int(rec.get("數量") or 1)
                for i in range(1, total_qty + 1):
                    img = make_label_image(rec, i, total_qty)
                    print_label_simple(printer, img)
                success += 1
                self._status.set(f"列印中... {success}/{total}")
                self.update()
            except Exception as e:
                messagebox.showerror("列印錯誤", f"SN {sn} 列印失敗：{e}")

        self._status.set(f"列印完成，共 {success} 筆")
        messagebox.showinfo("完成", f"列印完成，共列印 {success} 筆")

    def _parse_range(self, text, max_val):
        """解析張數輸入字串，回傳排序後的整數 list。
        支援格式：1,2,3  /  30-33  /  1-3,5,30-33
        """
        nums = set()
        for part in text.replace('，', ',').replace('~', '-').replace('到', '-').split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                a, _, b = part.partition('-')
                try:
                    a, b = int(a.strip()), int(b.strip())
                    nums.update(range(a, b + 1))
                except ValueError:
                    pass
            else:
                try:
                    nums.add(int(part))
                except ValueError:
                    pass
        return sorted(n for n in nums if 1 <= n <= max_val)

    def _print_range(self):
        if not self._selected_ids:
            messagebox.showwarning("提示", "請先勾選一筆要重印的資料")
            return
        if len(self._selected_ids) > 1:
            messagebox.showwarning("提示", "指定張數每次只能選一筆資料")
            return

        sn = next(iter(self._selected_ids))
        row = self._get_record_by_sn(sn)
        if not row:
            return
        rec = self._build_record_dict(row)
        total = int(rec.get("數量") or 1)
        printer = self._printer_var.get()
        if not printer:
            messagebox.showwarning("提示", "請選擇印表機")
            return

        PrintRangeDialog(self, rec, total, printer)

    def _clear_data(self):
        if not messagebox.askyesno("確認", "確定要清除所有資料嗎？\n此動作無法還原"):
            return
        conn = get_db()
        conn.execute("DELETE FROM MYTABLE")
        conn.commit()
        conn.close()
        self._selected_ids.clear()
        self._refresh_table()


# ── 預覽視窗 ──────────────────────────────────────────────────────────────────

class PrintRangeDialog(tk.Toplevel):
    """指定張數列印對話框"""

    def __init__(self, parent, rec, total, printer):
        super().__init__(parent)
        self.title("指定張數列印")
        self._rec = rec
        self._total = total
        self._printer = printer
        self._parent = parent

        # 先設定視窗大小與位置（避免 resizable(False) 造成內容不顯示）
        w, h = 560, 340
        px = parent.winfo_rootx() + parent.winfo_width() // 2 - w // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - h // 2
        self.geometry(f"{w}x{h}+{max(0,px)}+{max(0,py)}")
        self.minsize(w, h)

        self._build(rec, total)
        self.update_idletasks()
        self.transient(parent)
        self.lift()
        self.focus_force()
        self.after(50, self.grab_set)

    def _build(self, rec, total):
        pad = dict(padx=10, pady=4)

        # 資料摘要
        info = (f"供應商：{rec.get('供應商名稱','')}　"
                f"批號：{rec.get('批號','')}　"
                f"總數量：{total} 張（1 ~ {total}）")
        ttk.Label(self, text=info, foreground="#444").pack(padx=10, pady=(10,2), anchor="w")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10)

        # 輸入說明
        hint = "輸入要列印的張數（支援逗號與範圍）："
        ttk.Label(self, text=hint).pack(**pad, anchor="w")

        eg = "範例：  1,2,3,5   或   30-33   或   1-3,5,30-33"
        ttk.Label(self, text=eg, foreground="gray").pack(padx=10, anchor="w")

        self._entry = ttk.Entry(self, width=36, font=("微軟正黑體", 12))
        self._entry.pack(padx=10, pady=6, fill="x")
        self._entry.focus()

        # 快捷按鈕
        quick_frame = ttk.Frame(self)
        quick_frame.pack(padx=10, pady=2, fill="x")
        ttk.Label(quick_frame, text="快速選取：").pack(side="left")
        ttk.Button(quick_frame, text=f"全部 1-{total}",
                   command=lambda: self._set(f"1-{total}")).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="最後5張",
                   command=lambda: self._set(f"{max(1,total-4)}-{total}")).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="第1張",
                   command=lambda: self._set("1")).pack(side="left", padx=2)

        # 預覽張數
        self._preview_lbl = ttk.Label(self, text="", foreground="#0066cc")
        self._preview_lbl.pack(padx=10, pady=2, anchor="w")
        self._entry.bind("<KeyRelease>", self._on_key)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=4)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(padx=10, pady=(0,10))
        ttk.Button(btn_frame, text="列印", width=12, command=self._do_print).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="取消", width=12, command=self.destroy).pack(side="left", padx=4)

    def _set(self, text):
        self._entry.delete(0, tk.END)
        self._entry.insert(0, text)
        self._on_key()

    def _on_key(self, event=None):
        nums = self._parent._parse_range(self._entry.get(), self._total)
        if nums:
            self._preview_lbl.config(
                text=f"將列印 {len(nums)} 張：{', '.join(map(str, nums[:10]))}{'...' if len(nums)>10 else ''}",
                foreground="#0066cc")
        else:
            self._preview_lbl.config(text="（尚未輸入有效張數）", foreground="gray")

    def _do_print(self):
        nums = self._parent._parse_range(self._entry.get(), self._total)
        if not nums:
            messagebox.showwarning("提示", "請輸入有效的張數", parent=self)
            return
        if not messagebox.askyesno("確認列印",
                f"即將列印 {len(nums)} 張\n印表機：{self._printer}\n\n確定？", parent=self):
            return

        self.destroy()
        success = 0
        for i, n in enumerate(nums, 1):
            try:
                img = make_label_image(self._rec, n, self._total)
                print_label_simple(self._printer, img)
                success += 1
                self._parent._status.set(f"指定列印中... {i}/{len(nums)}")
                self._parent.update()
            except Exception as e:
                messagebox.showerror("列印錯誤", f"第 {n} 張失敗：{e}")
        self._parent._status.set(f"指定列印完成，共 {success} 張")
        messagebox.showinfo("完成", f"列印完成，共列印 {success} 張")


class PreviewWindow(tk.Toplevel):
    def __init__(self, parent, pil_image):
        super().__init__(parent)
        self.title("標籤預覽")
        self.resizable(False, False)

        # 縮放到螢幕適合的大小
        scale = 0.8
        w = int(pil_image.width * scale)
        h = int(pil_image.height * scale)
        preview = pil_image.resize((w, h), Image.LANCZOS)

        from PIL import ImageTk
        self._photo = ImageTk.PhotoImage(preview)
        lbl = tk.Label(self, image=self._photo, relief="solid", bd=1)
        lbl.pack(padx=10, pady=10)

        ttk.Label(self, text=f"實際尺寸：{LABEL_W_MM}mm × {LABEL_H_MM}mm  /  {LABEL_W_PX}px × {LABEL_H_PX}px @ {PRINT_DPI}DPI",
                  foreground="gray").pack(pady=2)
        ttk.Button(self, text="關閉", command=self.destroy).pack(pady=6)
        self.grab_set()


# ── 啟動 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    app = App()
    app.mainloop()
