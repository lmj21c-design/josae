# -*- coding: utf-8 -*-
"""
제비스코 충주점 라벨 프린터 v3
Brother QL-700 / 62mm 테이프 / 배민 한나체
"""

import struct
import json
import os
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont, ImageWin, ImageTk
import win32print
import win32ui
import win32con

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── 공통 색상 팔레트 (베이지/모던 화이트) ──────────────
UI_BG       = '#f5f3ee'   # 페이지 배경
UI_CARD     = '#ffffff'   # 카드 배경
UI_BORDER   = '#e3e0d8'   # 카드 테두리
UI_TEXT     = '#2b2924'   # 본문 텍스트
UI_MUTED    = '#9a9588'   # 보조 텍스트
UI_DARK_BTN = '#2b2924'   # 강조 버튼(상호명과 동일 톤)
UI_DARK_BTN_HOVER = '#46423a'
UI_ACCENT   = '#0f766e'   # 포인트(고객등록 등)
UI_ACCENT_HOVER = '#0d5f58'

def style_window(win):
    win.configure(fg_color=UI_BG)

def flat_button(parent, text, command, primary=False, accent=False, width=None):
    if primary:
        bg, hover, fg = UI_DARK_BTN, UI_DARK_BTN_HOVER, '#ffffff'
    elif accent:
        bg, hover, fg = UI_ACCENT, UI_ACCENT_HOVER, '#ffffff'
    else:
        bg, hover, fg = UI_CARD, '#f0eee7', UI_TEXT
    btn = ctk.CTkButton(parent, text=text, command=command,
                         font=("맑은 고딕", 13), fg_color=bg, hover_color=hover,
                         text_color=fg, corner_radius=10,
                         border_width=0 if (primary or accent) else 1,
                         border_color=UI_BORDER,
                         height=40)
    if width: btn.configure(width=width)
    return btn

# ── 설정 파일 ──────────────────────────────────────────
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jevisco_config.json')

DEFAULT_CONFIG = {
    "HISTORY_PATH" : r"C:\local\HISTORY.DBF",
    "ROOT_SU"      : r"C:\수성",
    "ROOT_YU"      : r"C:\유성",
    "FONT_PATH"    : r"C:\Windows\Fonts\BMHANNAPro.ttf",
    "SHOP_NAME"    : "페인트 제비스코 충주점",
    "SHOP_SLOGAN"  : "도·소매 / 공장납품전문 / 조색 / 상담환영",
    "SHOP_ADDR"    : "충북 충주시 봉현로 105",
    "SHOP_TEL"     : "TEL : 043) 911-4599",
    "SHOP_FAX"     : "FAX : 043) 911-4600",
    "CUSTOMER_PIN" : "1234",

    # ── 디자인 설정 (LAYOUT) ──
    "L_SHOP_NAME_SIZE"   : 36,
    "L_SHOP_NAME_COLOR"  : "#000000",
    "L_PRODUCT_SIZE"     : 26,
    "L_PRODUCT_COLOR"    : "#111111",
    "L_LABEL_SIZE"       : 18,
    "L_LABEL_COLOR"      : "#888888",
    "L_CODE_SIZE"        : 72,
    "L_CODE_FG_COLOR"    : "#ffffff",
    "L_CODE_BG_COLOR"    : "#000000",
    "L_SLOGAN_SIZE"      : 20,
    "L_SLOGAN_COLOR"     : "#333333",
    "L_DATETIME_SIZE"    : 20,
    "L_DATETIME_COLOR"   : "#111111",
    "L_CONTACT_SIZE"     : 18,
    "L_CONTACT_COLOR"    : "#444444",
    "L_BORDER_STYLE"     : "classic",   # classic / simple / boxed
    "L_DECORATION"       : "wave",      # wave / line / none
    "L_CODE_Y"           : 210,         # 조색번호 박스 세로 중심
}

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except: pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

CFG = load_config()

def get_cfg(key):
    val = CFG.get(key, DEFAULT_CONFIG.get(key, ''))
    default = DEFAULT_CONFIG.get(key)
    if isinstance(default, int) and not isinstance(val, int):
        try: return int(val)
        except: return default
    return val

# ── 고객 데이터 관리 ───────────────────────────────────
CUSTOMERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jevisco_customers.json')

def load_customers():
    if os.path.exists(CUSTOMERS_PATH):
        try:
            with open(CUSTOMERS_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return []

def save_customers(customers):
    with open(CUSTOMERS_PATH, 'w', encoding='utf-8') as f:
        json.dump(customers, f, ensure_ascii=False, indent=2)

def mask_phone(phone):
    """010-1234-5678 -> 010-****-5678"""
    if not phone:
        return ''
    digits_only = phone.replace('-', '')
    if len(digits_only) < 7:
        return phone
    parts = phone.split('-')
    if len(parts) == 3:
        return f"{parts[0]}-****-{parts[2]}"
    # 구분자가 없는 경우 중간 부분 마스킹
    n = len(digits_only)
    return digits_only[:3] + '*' * (n - 7) + digits_only[-4:]

def next_customer_id(customers):
    if not customers:
        return 1
    return max(c.get('id', 0) for c in customers) + 1

# ── DBF 읽기 ───────────────────────────────────────────
def read_dbf(path, encoding='cp949'):
    records = []
    fields = []
    with open(path, 'rb') as f:
        f.read(1); f.read(3)
        num_records = struct.unpack('<I', f.read(4))[0]
        f.read(2); f.read(2); f.read(20)
        while True:
            t = f.read(1)
            if t == b'\r': break
            name_raw = t + f.read(10)
            name = name_raw.rstrip(b'\x00').decode('ascii', errors='replace')
            ftype = f.read(1).decode('ascii')
            f.read(4)
            flen = struct.unpack('B', f.read(1))[0]
            f.read(1); f.read(14)
            fields.append((name, ftype, flen))
        for _ in range(num_records):
            f.read(1)
            row = {}
            for name, _, length in fields:
                raw = f.read(length)
                try: val = raw.decode(encoding).strip()
                except: val = raw.decode('latin1').strip()
                row[name] = val
            records.append(row)
    return records

def get_records(count=None):
    records = read_dbf(get_cfg('HISTORY_PATH'))
    if not records:
        raise ValueError("HISTORY.DBF에 데이터가 없습니다.")
    # DBF 필드 목록을 한 번만 저장 (필드 확인용)
    try:
        dbf_info_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dbf_fields.txt')
        if not os.path.exists(dbf_info_path):
            latest = records[-1]  # 가장 최근 레코드
            with open(dbf_info_path, 'w', encoding='utf-8') as f:
                f.write("=== HISTORY.DBF 필드 목록 ===\n")
                f.write(", ".join(latest.keys()) + "\n\n")
                f.write("=== 가장 최근 레코드 데이터 ===\n")
                for k, v in latest.items():
                    f.write(f"{k}: {v}\n")
    except: pass
    if count:
        return records[-count:][::-1]
    return records[::-1]

def lookup_product(prd_id, disp_id):
    root = get_cfg('ROOT_SU') if disp_id == '1' else get_cfg('ROOT_YU')
    path = os.path.join(root, 'PRODUCTS.DBF')
    try:
        for r in read_dbf(path):
            if r.get('ID','').strip() == str(prd_id).strip():
                return r.get('DESCR', r.get('CODE','')), r.get('PATH','')
    except: pass
    return f"제품ID:{prd_id}", ''

def lookup_base(bas_id, disp_id):
    root = get_cfg('ROOT_SU') if disp_id == '1' else get_cfg('ROOT_YU')
    path = os.path.join(root, 'BASES.DBF')
    try:
        for r in read_dbf(path):
            if r.get('ID','').strip() == str(bas_id).strip():
                return r.get('DESCR', r.get('CODE',''))
    except: pass
    return ''

GLOSS_KEYWORDS = ['유광', '무광', '반광', '에그쉘광']

def lookup_gloss(spd_id, prd_path, disp_id):
    if not spd_id or not prd_path: return '', ''
    root = get_cfg('ROOT_SU') if disp_id == '1' else get_cfg('ROOT_YU')
    sub_path = os.path.join(root, prd_path, 'SUBPRODS.DBF')
    try:
        for r in read_dbf(sub_path):
            if r.get('ID','').strip() == str(spd_id).strip():
                descr = r.get('DESCR', r.get('CODE','')).strip()
                if any(k in descr for k in GLOSS_KEYWORDS):
                    return descr, ''
                else:
                    return '', descr
    except: pass
    return '', ''

# ── 날짜/시간 포맷 ─────────────────────────────────────
WEEKDAYS = ['월요일','화요일','수요일','목요일','금요일','토요일','일요일']

def format_date(date_str):
    try:
        dt = datetime.strptime(date_str.strip(), "%Y-%m-%d")
        return f"{dt.year}년 {dt.month}월 {dt.day}일 {WEEKDAYS[dt.weekday()]}"
    except: return date_str

def format_time(time_str):
    try:
        dt = datetime.strptime(time_str.strip(), "%H:%M:%S")
        ampm = "오전" if dt.hour < 12 else "오후"
        hour = dt.hour if dt.hour <= 12 else dt.hour - 12
        return f"{ampm} {hour}:{dt.minute:02d}"
    except: return time_str

def format_volume(nom_q_str):
    """NOM_Q(g) → 용량 문자열. 알려진 근사값만 표시, 나머지는 빈 문자열"""
    try:
        grams = float(nom_q_str.strip())
    except:
        return ''
    # (최솟값, 최댓값, 표시문자열) 순으로 근사 매핑
    mapping = [
        (900,  1050,  '1L'),
        (3900, 4100,  '4L'),
        (3200, 3500,  '3.6L'),
        (18000, 18500, '18L'),
        (14500, 15000, '16L'),
    ]
    for lo, hi, label in mapping:
        if lo <= grams <= hi:
            return label
    return ''  # 알 수 없는 용량은 표시 안 함

# ── 레코드 파싱 ────────────────────────────────────────
def parse_record(rec):
    disp_id  = rec.get('DISP_ID','1').strip()
    prd_id   = rec.get('PRDID','').strip()
    bas_id   = rec.get('BASID','').strip()
    spd_id   = rec.get('SPDID','').strip()
    color_code = rec.get('CLRDESCR','').split('|')[0].strip()
    date_str = rec.get('DATA','').strip()
    time_str = rec.get('TIME','').strip()
    volume   = format_volume(rec.get('NOM_Q',''))

    disp_type = '수성' if disp_id == '1' else '유성'
    product_name, prd_path = lookup_product(prd_id, disp_id)
    base_name  = lookup_base(bas_id, disp_id)
    gloss, override_name = lookup_gloss(spd_id, prd_path, disp_id)
    if override_name:
        product_name = override_name

    return {
        'product_name': product_name,
        'base_name'   : base_name,
        'gloss'       : gloss,
        'color_code'  : color_code,
        'disp_type'   : disp_type,
        'date_str'    : date_str,
        'time_str'    : time_str,
        'volume'      : volume,
    }

# ── 라벨 이미지 생성 ────────────────────────────────────
W, H = 732, 410

def fnt(size):
    return ImageFont.truetype(get_cfg('FONT_PATH'), size)


def draw_wave_line(dr, x1, x2, y, amplitude=4, segments=20, width=2):
    import math
    step = (x2 - x1) / segments
    pts = [(x1 + i * step, y + amplitude * math.sin(math.pi * i / (segments / 2)))
           for i in range(segments + 1)]
    for i in range(len(pts) - 1):
        dr.line([pts[i], pts[i+1]], fill='black', width=width)

def draw_label(d):
    gloss      = d['gloss']
    color_code = d['color_code']
    disp_type  = d['disp_type']
    date_fmt   = format_date(d['date_str'])
    time_fmt   = format_time(d['time_str'])
    prd_line   = d['product_name']
    if gloss: prd_line += f"  {gloss}"
    prd_line  += f"  ( {disp_type} )"

    style = get_cfg('L_BORDER_STYLE')
    deco  = get_cfg('L_DECORATION')

    img = Image.new('RGB', (W, H), 'white')
    dr  = ImageDraw.Draw(img)

    # ── 외곽 테두리 ──
    if style == 'classic':
        dr.rectangle([4,  4,  W-5,  H-5],  outline='black', width=7)
        dr.rectangle([12, 12, W-13, H-13], outline='black', width=2)
        dr.rectangle([16, 16, W-17, H-17], outline='black', width=1)
        for pts in [
            [(4,4),(34,4),(4,34)], [(W-4,4),(W-34,4),(W-4,34)],
            [(4,H-4),(34,H-4),(4,H-34)], [(W-4,H-4),(W-34,H-4),(W-4,H-34)],
        ]:
            dr.polygon(pts, fill='black')
    elif style == 'boxed':
        dr.rectangle([4, 4, W-5, H-5], outline='black', width=10)
    else:  # simple
        dr.rectangle([6, 6, W-7, H-7], outline='black', width=3)

    def divider(y, double=True):
        if double:
            dr.line([16, y-2, W-16, y-2], fill='black', width=1)
            dr.line([16, y,   W-16, y],   fill='black', width=3)
        else:
            dr.line([16, y, W-16, y], fill='black', width=2)
        if deco == 'wave':
            draw_wave_line(dr, 60, W-60, y-8, amplitude=4, segments=20, width=2)

    # ── 상호명 ──
    dr.line([16, 28, W-16, 28], fill='black', width=1)
    divider(80)
    dr.line([16, 84, W-16, 84], fill='black', width=1)
    dr.text((W//2, 56), get_cfg('SHOP_NAME'),
             font=fnt(get_cfg('L_SHOP_NAME_SIZE')), fill=get_cfg('L_SHOP_NAME_COLOR'), anchor='mm')

    # ── 제품명 ──
    dr.text((W//2, 108), prd_line,
             font=fnt(get_cfg('L_PRODUCT_SIZE')), fill=get_cfg('L_PRODUCT_COLOR'), anchor='mm')
    divider(126)
    dr.line([16, 130, W-16, 130], fill='black', width=1)

    # ── 조색번호 레이블 ──
    dr.text((W//2, 150), "조  색  번  호",
             font=fnt(get_cfg('L_LABEL_SIZE')), fill=get_cfg('L_LABEL_COLOR'), anchor='mm')

    # ── 조색번호 박스 ──
    code_y = get_cfg('L_CODE_Y')
    box_top, box_bottom = 158, 262
    dr.rectangle([60, box_top, W-60, box_bottom], fill=get_cfg('L_CODE_BG_COLOR'))
    dr.text((W//2, code_y), color_code,
             font=fnt(get_cfg('L_CODE_SIZE')), fill=get_cfg('L_CODE_FG_COLOR'), anchor='mm')

    # ── 하단 구분선 ──
    divider(276)
    dr.line([16, 273, W-16, 273], fill='black', width=1)

    # ── 슬로건 ──
    dr.text((W//2, 302), get_cfg('SHOP_SLOGAN'),
             font=fnt(get_cfg('L_SLOGAN_SIZE')), fill=get_cfg('L_SLOGAN_COLOR'), anchor='mm')
    dr.line([16, 316, W-16, 316], fill='#bbbbbb', width=1)

    # ── 날짜/시간 ──
    dt_font  = fnt(get_cfg('L_DATETIME_SIZE'))
    dt_color = get_cfg('L_DATETIME_COLOR')
    dr.text((44, 334), f"조색날짜 : {date_fmt}", font=dt_font, fill=dt_color)
    dr.text((44, 364), f"조색시간 : {time_fmt}", font=dt_font, fill=dt_color)

    # ── 세로 구분선 + 주소 ──
    dr.line([496, 318, 496, H-14], fill='#cccccc', width=1)
    ct_font  = fnt(get_cfg('L_CONTACT_SIZE'))
    ct_color = get_cfg('L_CONTACT_COLOR')
    dr.text((670, 326), get_cfg('SHOP_ADDR'), font=ct_font, fill=ct_color, anchor='ra')
    dr.text((670, 348), get_cfg('SHOP_TEL'),  font=ct_font, fill=ct_color, anchor='ra')
    dr.text((670, 370), get_cfg('SHOP_FAX'),  font=ct_font, fill=ct_color, anchor='ra')

    return img

# ── 프린터 출력 ────────────────────────────────────────
def print_to_brother(img, copies=1):
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    printer_name = win32print.GetDefaultPrinter()
    for p in printers:
        if 'QL-700' in p or 'Brother' in p:
            printer_name = p
            break
    for _ in range(copies):
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        pw = hDC.GetDeviceCaps(win32con.HORZRES)
        ph = hDC.GetDeviceCaps(win32con.VERTRES)
        img_r = img.resize((pw, ph), Image.LANCZOS)
        hDC.StartDoc("제비스코 라벨")
        hDC.StartPage()
        dib = ImageWin.Dib(img_r)
        dib.draw(hDC.GetHandleOutput(), (0, 0, pw, ph))
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

# ── 미리보기 + 출력 팝업 ───────────────────────────────
class PrintPreviewDialog(ctk.CTkToplevel):
    def __init__(self, parent, data):
        super().__init__(parent)
        self.title("라벨 미리보기")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)
        self.data = data
        self.result = False

        img = draw_label(data)
        # 미리보기용 축소 (가로 500px)
        ratio = 500 / img.width
        preview = img.resize((500, int(img.height * ratio)), Image.LANCZOS)
        self.ctk_img = ctk.CTkImage(light_image=preview, size=preview.size)

        ctk.CTkLabel(self, image=self.ctk_img, text="", fg_color=UI_BG).pack(padx=16, pady=(16,8))

        # 출력 수량
        qty_frame = ctk.CTkFrame(self, fg_color=UI_BG)
        qty_frame.pack(pady=(0,12))
        ctk.CTkLabel(qty_frame, text="출력 수량 :", font=("맑은 고딕", 12),
                     fg_color=UI_BG, text_color=UI_TEXT).pack(side="left", padx=(0,8))
        self.qty_var = tk.StringVar(value="1")
        ctk.CTkEntry(qty_frame, textvariable=self.qty_var, width=50,
                     font=("맑은 고딕", 12)).pack(side="left")
        ctk.CTkLabel(qty_frame, text="장", font=("맑은 고딕", 12),
                     fg_color=UI_BG, text_color=UI_TEXT).pack(side="left", padx=(4,0))

        btn_frame = ctk.CTkFrame(self, fg_color=UI_BG)
        btn_frame.pack(pady=(0,16))
        flat_button(btn_frame, "출력", self.on_print, primary=True).pack(side="left", padx=8)
        flat_button(btn_frame, "취소", self.destroy).pack(side="left", padx=8)

    def on_print(self):
        try:
            copies = int(self.qty_var.get())
        except: copies = 1
        img = draw_label(self.data)
        print_to_brother(img, copies)
        self.result = True
        self.destroy()

# ── 이전 목록 팝업 ─────────────────────────────────────
class HistoryDialog(ctk.CTkToplevel):
    def __init__(self, parent, multi_select=False):
        super().__init__(parent)
        self.title("조색 목록")
        self.geometry("660x420")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)
        self.selected = None
        self.selected_multi = []
        self.multi_select = multi_select

        # 검색 영역
        search_frame = ctk.CTkFrame(self, fg_color=UI_BG)
        search_frame.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(search_frame, text="날짜:", font=("맑은 고딕", 11),
                     fg_color=UI_BG, text_color=UI_TEXT).pack(side="left")
        self.date_var = tk.StringVar()
        ctk.CTkEntry(search_frame, textvariable=self.date_var, width=110,
                     font=("맑은 고딕", 11)).pack(side="left", padx=(4,12))

        ctk.CTkLabel(search_frame, text="조색번호:", font=("맑은 고딕", 11),
                     fg_color=UI_BG, text_color=UI_TEXT).pack(side="left")
        self.code_var = tk.StringVar()
        ctk.CTkEntry(search_frame, textvariable=self.code_var, width=110,
                     font=("맑은 고딕", 11)).pack(side="left", padx=(4,12))

        flat_button(search_frame, "검색", self.do_search).pack(side="left", padx=4)
        flat_button(search_frame, "전체", self.load_all).pack(side="left", padx=4)

        if multi_select:
            ctk.CTkLabel(self, text="Ctrl 또는 Shift를 눌러 여러 항목을 선택할 수 있습니다",
                         font=("맑은 고딕", 10), fg_color=UI_BG, text_color=UI_MUTED).pack(
                             anchor='w', padx=14, pady=(0,4))

        # 리스트박스
        list_frame = ctk.CTkFrame(self, fg_color=UI_CARD, corner_radius=10,
                                   border_width=1, border_color=UI_BORDER)
        list_frame.pack(fill="both", expand=True, padx=12)
        list_inner = tk.Frame(list_frame, bg=UI_CARD)
        list_inner.pack(fill="both", expand=True, padx=8, pady=8)

        scrollbar = tk.Scrollbar(list_inner)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(list_inner, font=("맑은 고딕", 10),
                                   yscrollcommand=scrollbar.set,
                                   selectmode=(tk.EXTENDED if multi_select else tk.SINGLE),
                                   bg=UI_CARD, fg=UI_TEXT,
                                   highlightthickness=0, relief='flat',
                                   selectbackground='#e3e0d8', selectforeground=UI_TEXT)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        if not multi_select:
            self.listbox.bind('<Double-Button-1>', lambda e: self.on_select())

        # 버튼
        btn_frame = ctk.CTkFrame(self, fg_color=UI_BG)
        btn_frame.pack(pady=10)
        btn_label = "선택 연결" if multi_select else "출력"
        flat_button(btn_frame, btn_label, self.on_select, primary=True).pack(side="left", padx=8)
        flat_button(btn_frame, "취소", self.destroy).pack(side="left", padx=8)

        self.all_records = []
        self.filtered = []
        self.load_all()

    def load_all(self):
        self.date_var.set('')
        self.code_var.set('')
        self.all_records = get_records(100)
        self._populate(self.all_records)

    def do_search(self):
        date_q = self.date_var.get().strip()
        code_q = self.code_var.get().strip().upper()
        result = []
        for rec in self.all_records:
            date_str   = rec.get('DATA','').strip()
            color_code = rec.get('CLRDESCR','').split('|')[0].strip().upper()
            if date_q and date_q not in date_str: continue
            if code_q and code_q not in color_code: continue
            result.append(rec)
        self._populate(result)

    def _populate(self, records):
        self.filtered = records
        self.listbox.delete(0, tk.END)
        for rec in records:
            disp_id    = rec.get('DISP_ID','1').strip()
            prd_id     = rec.get('PRDID','').strip()
            spd_id     = rec.get('SPDID','').strip()
            color_code = rec.get('CLRDESCR','').split('|')[0].strip()
            date_str   = rec.get('DATA','').strip()
            time_str   = rec.get('TIME','').strip()[:5]
            disp_type  = '수성' if disp_id == '1' else '유성'
            prd_name, prd_path = lookup_product(prd_id, disp_id)
            gloss, override_name = lookup_gloss(spd_id, prd_path, disp_id)
            if override_name:
                prd_name = override_name
            gloss_disp = f" {gloss}" if gloss else ""
            volume     = format_volume(rec.get('NOM_Q',''))
            vol_disp   = f" [{volume}]" if volume else ""
            self.listbox.insert(tk.END,
                f"{date_str} {time_str}  |  {prd_name}{gloss_disp}{vol_disp} ({disp_type})  |  {color_code}")

    def on_select(self):
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("선택 없음", "항목을 선택해주세요.", parent=self)
            return
        if self.multi_select:
            self.selected_multi = [self.filtered[i] for i in sel]
        else:
            self.selected = self.filtered[sel[0]]
        self.destroy()

# ── 고객 등록/조회 화면 ─────────────────────────────────
class CustomerDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("고객 등록 / 조회")
        self.geometry("980x740")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)

        self.customers = load_customers()
        self.filtered = list(self.customers)
        self.current_customer = None
        self.phone_revealed = False

        # ── 검색 영역 (카드) ──
        search_card = tk.Frame(self, bg=UI_CARD, highlightthickness=1,
                                highlightbackground=UI_BORDER)
        search_card.pack(fill=tk.X, padx=14, pady=(14,10))
        search_frame = tk.Frame(search_card, bg=UI_CARD, pady=10, padx=10)
        search_frame.pack(fill=tk.X)
        tk.Label(search_frame, text="이름 검색:", font=("맑은 고딕", 10),
                 bg=UI_CARD, fg=UI_TEXT).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        entry = tk.Entry(search_frame, textvariable=self.search_var, width=18,
                          font=("맑은 고딕", 10))
        entry.pack(side=tk.LEFT, padx=(4,8))
        entry.bind('<Return>', lambda e: self.do_search())
        flat_button(search_frame, "검색", self.do_search, primary=True).pack(side=tk.LEFT, padx=2)
        flat_button(search_frame, "전체", self.load_all).pack(side=tk.LEFT, padx=2)

        # ── 좌/우 분할 ──
        body = tk.Frame(self, bg=UI_BG)
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0,10))

        # 좌측: 고객 목록 (카드)
        left = tk.Frame(body, bg=UI_CARD, highlightthickness=1,
                         highlightbackground=UI_BORDER)
        left.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0,8))
        left_inner = tk.Frame(left, bg=UI_CARD, padx=12, pady=12)
        left_inner.pack(fill=tk.BOTH, expand=True)
        tk.Label(left_inner, text="고객 목록", font=("맑은 고딕", 10, "bold"),
                 bg=UI_CARD, fg=UI_MUTED, anchor='w').pack(fill=tk.X, pady=(0,8))

        list_frame = tk.Frame(left_inner, bg=UI_CARD)
        list_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox = tk.Listbox(list_frame, font=("맑은 고딕", 11), width=30, height=18,
                                   yscrollcommand=scrollbar.set, selectmode=tk.SINGLE,
                                   bg=UI_CARD, fg=UI_TEXT, highlightthickness=0, relief='flat',
                                   selectbackground=UI_BG, selectforeground=UI_TEXT)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind('<<ListboxSelect>>', lambda e: self.on_customer_select())

        flat_button(left_inner, "+ 신규 고객 등록", self.open_new_customer).pack(fill=tk.X, pady=(10,0))

        # 우측: 상세 정보 (카드)
        right = tk.Frame(body, width=600, bg=UI_CARD, highlightthickness=1,
                          highlightbackground=UI_BORDER)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_inner = tk.Frame(right, bg=UI_CARD, padx=14, pady=14)
        right_inner.pack(fill=tk.BOTH, expand=True)

        name_row = tk.Frame(right_inner, bg=UI_CARD)
        name_row.pack(fill=tk.X, pady=(0,2))
        self.detail_name = tk.Label(name_row, text="고객을 선택하세요",
                                     font=("맑은 고딕", 16, "bold"),
                                     bg=UI_CARD, fg=UI_TEXT, anchor='w')
        self.detail_name.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.edit_btn = tk.Button(name_row, text="수정", font=("맑은 고딕", 10),
                                   relief='flat', bd=0, fg=UI_ACCENT, bg=UI_CARD,
                                   activeforeground=UI_ACCENT_HOVER, activebackground=UI_CARD,
                                   cursor='hand2', command=self.open_edit_customer)
        self.delete_btn = tk.Button(name_row, text="삭제", font=("맑은 고딕", 10),
                                     relief='flat', bd=0, fg='#c0392b', bg=UI_CARD,
                                     activeforeground='#992d22', activebackground=UI_CARD,
                                     cursor='hand2', command=self.delete_customer)

        phone_frame = tk.Frame(right_inner, bg=UI_CARD)
        phone_frame.pack(fill=tk.X, pady=(0,12))
        self.detail_phone = tk.Label(phone_frame, text="", font=("맑은 고딕", 12),
                                      bg=UI_CARD, fg=UI_MUTED, anchor='w')
        self.detail_phone.pack(side=tk.LEFT)
        self.reveal_btn = tk.Button(phone_frame, text="전체보기", font=("맑은 고딕", 10, "underline"),
                                     relief='flat', bd=0, fg=UI_ACCENT, bg=UI_CARD,
                                     activeforeground=UI_ACCENT_HOVER, activebackground=UI_CARD,
                                     cursor='hand2', command=self.on_reveal_phone)
        self.reveal_btn.pack(side=tk.LEFT, padx=(8,0))

        tk.Label(right_inner, text="연결된 조색이력", font=("맑은 고딕", 12, "bold"),
                 bg=UI_CARD, fg=UI_MUTED, anchor='w').pack(fill=tk.X, pady=(4,6))

        hist_frame = tk.Frame(right_inner, bg=UI_CARD)
        hist_frame.pack(fill=tk.BOTH, expand=True)
        hscroll = tk.Scrollbar(hist_frame)
        hscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_listbox = tk.Listbox(hist_frame, font=("맑은 고딕", 11), height=16,
                                           yscrollcommand=hscroll.set, selectmode=tk.EXTENDED,
                                           bg=UI_CARD, fg=UI_TEXT, highlightthickness=0,
                                           relief='flat', selectbackground=UI_BG,
                                           selectforeground=UI_TEXT)
        self.history_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hscroll.config(command=self.history_listbox.yview)

        link_btn_row = tk.Frame(right_inner, bg=UI_CARD)
        link_btn_row.pack(fill=tk.X, pady=(10,0))
        flat_button(link_btn_row, "+ 조색이력 연결하기", self.open_link_history).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        flat_button(link_btn_row, "선택 항목 삭제", self.delete_selected_history).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(4,0))

        # ── 하단 버튼 ──
        btn_frame = tk.Frame(self, bg=UI_BG)
        btn_frame.pack(pady=(0,14))
        flat_button(btn_frame, "닫기", self.destroy).pack(side=tk.LEFT, padx=8)

        self.load_all()

    def load_all(self):
        self.search_var.set('')
        self.customers = load_customers()
        self.filtered = list(self.customers)
        self._populate()

    def do_search(self):
        q = self.search_var.get().strip()
        if not q:
            self.filtered = list(self.customers)
        else:
            self.filtered = [c for c in self.customers if q in c.get('name','')]
        self._populate()

    def _populate(self):
        self.listbox.delete(0, tk.END)
        for c in self.filtered:
            phone = mask_phone(c.get('phone',''))
            phone_disp = phone if phone else '연락처 없음'
            self.listbox.insert(tk.END, f"{c.get('name','')}   {phone_disp}")
        self.current_customer = None
        self.phone_revealed = False
        self.detail_name.config(text="고객을 선택하세요")
        self.detail_phone.config(text="")
        self.edit_btn.pack_forget()
        self.delete_btn.pack_forget()
        self.history_listbox.delete(0, tk.END)

    def on_customer_select(self):
        sel = self.listbox.curselection()
        if not sel: return
        self.current_customer = self.filtered[sel[0]]
        self.phone_revealed = False
        self._refresh_detail()

    def _refresh_detail(self):
        c = self.current_customer
        if not c: return
        self.detail_name.config(text=c.get('name',''))
        self.edit_btn.pack(side=tk.RIGHT, padx=(4,0))
        self.delete_btn.pack(side=tk.RIGHT, padx=(4,0))
        phone = c.get('phone','')
        if not phone:
            self.detail_phone.config(text="연락처 없음")
            self.reveal_btn.pack_forget()
        else:
            shown = phone if self.phone_revealed else mask_phone(phone)
            self.detail_phone.config(text=shown)
            self.reveal_btn.config(text="가리기" if self.phone_revealed else "전체보기")
            self.reveal_btn.pack(side=tk.LEFT, padx=(8,0))

        self.history_listbox.delete(0, tk.END)
        records = sorted(c.get('history', []),
                          key=lambda r: (r.get('date',''), r.get('time','')), reverse=True)
        self._history_sorted = records
        if not records:
            self.history_listbox.insert(tk.END, "  (연결된 조색이력이 없습니다)")
        else:
            for h in records:
                date_disp = h.get('date','')
                time_disp = h.get('time','')[:5]
                gloss_disp = f" {h.get('gloss')}" if h.get('gloss') else ""
                vol_disp   = f" [{h.get('volume')}]" if h.get('volume') else ""
                self.history_listbox.insert(tk.END,
                    f"{h.get('color_code','')}   |   {h.get('product','')}{gloss_disp}{vol_disp}   |   {date_disp} {time_disp}")

    def on_reveal_phone(self):
        if self.phone_revealed:
            self.phone_revealed = False
            self._refresh_detail()
            return
        pin = get_cfg('CUSTOMER_PIN')
        entered = self._ask_pin()
        if entered is None:
            return
        if entered == pin:
            self.phone_revealed = True
            self._refresh_detail()
        else:
            messagebox.showerror("비밀번호 오류", "비밀번호가 일치하지 않습니다.", parent=self)

    def _ask_pin(self):
        dlg = tk.Toplevel(self)
        dlg.title("비밀번호 확인")
        dlg.resizable(False, False)
        dlg.attributes('-topmost', True)
        dlg.configure(bg=UI_BG)
        dlg.grab_set()
        result = {'value': None}

        tk.Label(dlg, text="고객 연락처를 보려면 비밀번호를 입력하세요",
                 font=("맑은 고딕", 10), bg=UI_BG, fg=UI_TEXT).pack(padx=20, pady=(16,8))
        var = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=var, show='*', font=("맑은 고딕", 12), width=16)
        entry.pack(pady=(0,12))
        entry.focus_set()

        def confirm():
            result['value'] = var.get()
            dlg.destroy()
        def cancel():
            dlg.destroy()

        entry.bind('<Return>', lambda e: confirm())
        btns = tk.Frame(dlg, bg=UI_BG)
        btns.pack(pady=(0,14))
        flat_button(btns, "확인", confirm, primary=True).pack(side=tk.LEFT, padx=6)
        flat_button(btns, "취소", cancel).pack(side=tk.LEFT, padx=6)

        self.wait_window(dlg)
        return result['value']

    def open_new_customer(self):
        dlg = CustomerEditDialog(self, None)
        self.wait_window(dlg)
        if dlg.saved:
            self.customers = load_customers()
            self.filtered = list(self.customers)
            self._populate()

    def open_link_history(self):
        if not self.current_customer:
            messagebox.showwarning("고객 미선택", "먼저 고객을 선택해주세요.", parent=self)
            return
        try:
            hist_dlg = HistoryDialog(self, multi_select=True)
            self.wait_window(hist_dlg)
            if hist_dlg.selected_multi:
                new_entries = []
                for rec in hist_dlg.selected_multi:
                    data = parse_record(rec)
                    new_entries.append({
                        'date': data['date_str'], 'time': data['time_str'],
                        'product': data['product_name'], 'color_code': data['color_code'],
                        'gloss': data.get('gloss',''), 'volume': data.get('volume',''),
                    })
                cust_list = load_customers()
                for c in cust_list:
                    if c['id'] == self.current_customer['id']:
                        c.setdefault('history', []).extend(new_entries)
                        break
                save_customers(cust_list)
                self.customers = cust_list
                for c in self.filtered:
                    if c['id'] == self.current_customer['id']:
                        c.setdefault('history', []).extend(new_entries)
                        self.current_customer = c
                        break
                self._refresh_detail()
        except Exception as e:
            messagebox.showerror("오류", str(e), parent=self)

    def open_edit_customer(self):
        if not self.current_customer:
            return
        dlg = CustomerEditDialog(self, self.current_customer)
        self.wait_window(dlg)
        if dlg.saved:
            self.customers = load_customers()
            updated = next((c for c in self.customers if c['id'] == self.current_customer['id']), None)
            self.do_search() if self.search_var.get().strip() else self.load_all()
            if updated:
                self.current_customer = updated
                for i, c in enumerate(self.filtered):
                    if c['id'] == updated['id']:
                        self.listbox.selection_clear(0, tk.END)
                        self.listbox.selection_set(i)
                        break
                self._refresh_detail()

    def delete_customer(self):
        if not self.current_customer:
            return
        name = self.current_customer.get('name', '')
        if not messagebox.askyesno("고객 삭제",
                f"'{name}' 고객 정보를 삭제하시겠습니까?\n연결된 조색이력도 함께 삭제됩니다.",
                parent=self):
            return
        customers = load_customers()
        customers = [c for c in customers if c['id'] != self.current_customer['id']]
        save_customers(customers)
        self.customers = customers
        self.filtered = [c for c in self.filtered if c['id'] != self.current_customer['id']]
        self._populate_keep_filter()

    def delete_selected_history(self):
        if not self.current_customer:
            messagebox.showwarning("고객 미선택", "먼저 고객을 선택해주세요.", parent=self)
            return
        sel = self.history_listbox.curselection()
        if not sel or not getattr(self, '_history_sorted', None):
            messagebox.showwarning("선택 없음", "삭제할 조색이력을 선택해주세요.", parent=self)
            return
        to_remove = [self._history_sorted[i] for i in sel]
        if not messagebox.askyesno("조색이력 연결 해제",
                f"선택한 {len(to_remove)}건의 조색이력 연결을 해제하시겠습니까?",
                parent=self):
            return

        cust_list = load_customers()
        for c in cust_list:
            if c['id'] == self.current_customer['id']:
                remaining = []
                for h in c.get('history', []):
                    if h in to_remove:
                        to_remove.remove(h)
                        continue
                    remaining.append(h)
                c['history'] = remaining
                self.current_customer = c
                break
        save_customers(cust_list)
        self.customers = cust_list
        for i, c in enumerate(self.filtered):
            if c['id'] == self.current_customer['id']:
                self.filtered[i] = self.current_customer
                break
        self._refresh_detail()

    def _populate_keep_filter(self):
        """삭제 후 검색 상태 유지하며 목록 갱신"""
        self.listbox.delete(0, tk.END)
        for c in self.filtered:
            phone = mask_phone(c.get('phone',''))
            phone_disp = phone if phone else '연락처 없음'
            self.listbox.insert(tk.END, f"{c.get('name','')}   {phone_disp}")
        self.current_customer = None
        self.phone_revealed = False
        self.detail_name.config(text="고객을 선택하세요")
        self.detail_phone.config(text="")
        self.edit_btn.pack_forget()
        self.delete_btn.pack_forget()
        self.history_listbox.delete(0, tk.END)


class CustomerEditDialog(ctk.CTkToplevel):
    def __init__(self, parent, customer):
        super().__init__(parent)
        self.is_edit = customer is not None
        self.title("고객 정보 수정" if self.is_edit else "신규 고객 등록")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)
        self.saved = False
        self.customer = customer

        tk.Label(self, text="이름 *", font=("맑은 고딕", 10), bg=UI_BG, fg=UI_TEXT).grid(
            row=0, column=0, padx=(16,8), pady=(16,4), sticky='w')
        self.name_var = tk.StringVar(value=(customer or {}).get('name',''))
        tk.Entry(self, textvariable=self.name_var, width=26,
                 font=("맑은 고딕", 10)).grid(row=0, column=1, padx=(0,16), pady=(16,4))

        tk.Label(self, text="연락처", font=("맑은 고딕", 10), bg=UI_BG, fg=UI_TEXT).grid(
            row=1, column=0, padx=(16,8), pady=4, sticky='w')
        self.phone_var = tk.StringVar(value=(customer or {}).get('phone',''))
        tk.Entry(self, textvariable=self.phone_var, width=26,
                 font=("맑은 고딕", 10)).grid(row=1, column=1, padx=(0,16), pady=4)

        # 최근 조색이력 선택 (신규 등록시에만 표시, 선택 입력)
        self.recent_records = []
        if not self.is_edit:
            tk.Label(self, text="최근 조색이력 연결 (선택)", font=("맑은 고딕", 10),
                      bg=UI_BG, fg=UI_TEXT, anchor='w').grid(
                          row=2, column=0, columnspan=2, padx=16, pady=(10,2), sticky='w')

            list_frame = tk.Frame(self, bg=UI_CARD, highlightthickness=1,
                                   highlightbackground=UI_BORDER)
            list_frame.grid(row=3, column=0, columnspan=2, padx=16, sticky='ew')
            scrollbar = tk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.hist_listbox = tk.Listbox(list_frame, font=("맑은 고딕", 9), height=8,
                                            selectmode=tk.MULTIPLE,
                                            yscrollcommand=scrollbar.set,
                                            bg=UI_CARD, fg=UI_TEXT, highlightthickness=0,
                                            relief='flat', selectbackground=UI_BG,
                                            selectforeground=UI_TEXT)
            self.hist_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.hist_listbox.yview)

            try:
                self.recent_records = get_records(50)
            except Exception:
                pass
            for rec in self.recent_records:
                date_str   = rec.get('DATA','').strip()
                time_str   = rec.get('TIME','').strip()[:5]
                color_code = rec.get('CLRDESCR','').split('|')[0].strip()
                disp_id    = rec.get('DISP_ID','1').strip()
                prd_id     = rec.get('PRDID','').strip()
                spd_id     = rec.get('SPDID','').strip()
                prd_name, prd_path = lookup_product(prd_id, disp_id)
                gloss, override_name = lookup_gloss(spd_id, prd_path, disp_id)
                if override_name:
                    prd_name = override_name
                gloss_disp = f" {gloss}" if gloss else ""
                self.hist_listbox.insert(tk.END,
                    f"{color_code}   |   {prd_name}{gloss_disp}   |   {date_str} {time_str}")

        btn_frame = tk.Frame(self, bg=UI_BG)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=14)
        flat_button(btn_frame, "저장", self.on_save, primary=True).pack(side=tk.LEFT, padx=8)
        flat_button(btn_frame, "취소", self.destroy).pack(side=tk.LEFT, padx=8)

    def on_save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("이름 필요", "이름은 필수 입력 항목입니다.", parent=self)
            return

        customers = load_customers()

        if self.is_edit:
            for c in customers:
                if c['id'] == self.customer['id']:
                    c['name'] = name
                    c['phone'] = self.phone_var.get().strip()
                    break
            save_customers(customers)
            self.saved = True
            messagebox.showinfo("수정 완료", f"{name} 고객님 정보가 수정되었습니다.", parent=self)
            self.destroy()
            return

        history_entries = []
        for idx in self.hist_listbox.curselection():
            rec = self.recent_records[idx]
            data = parse_record(rec)
            history_entries.append({
                'date': data['date_str'], 'time': data['time_str'],
                'product': data['product_name'], 'color_code': data['color_code'],
                'gloss': data.get('gloss',''), 'volume': data.get('volume',''),
            })

        new_id = next_customer_id(customers)
        customers.append({
            'id': new_id,
            'name': name,
            'phone': self.phone_var.get().strip(),
            'history': history_entries,
        })
        save_customers(customers)
        self.saved = True
        messagebox.showinfo("등록 완료", f"{name} 고객님이 등록되었습니다.", parent=self)
        self.destroy()

# ── 디자인 설정 화면 ───────────────────────────────────
class DesignSettingsDialog(ctk.CTkToplevel):
    SAMPLE_DATA = {
        'product_name': '7200 에나멜 프라임', 'base_name': '', 'gloss': '반광',
        'color_code': 'RAL7035', 'disp_type': '유성',
        'date_str': '2026-06-26', 'time_str': '14:27:00',
    }

    SLIDERS = [
        ("상호명 글씨 크기", "L_SHOP_NAME_SIZE", 16, 50),
        ("제품명 글씨 크기", "L_PRODUCT_SIZE",   12, 40),
        ("조색번호 레이블 크기", "L_LABEL_SIZE",  10, 30),
        ("조색번호 글씨 크기", "L_CODE_SIZE",     30, 100),
        ("조색번호 세로 위치", "L_CODE_Y",        180, 240),
        ("슬로건 글씨 크기", "L_SLOGAN_SIZE",     12, 32),
        ("날짜/시간 글씨 크기", "L_DATETIME_SIZE", 12, 30),
        ("주소/연락처 글씨 크기", "L_CONTACT_SIZE", 10, 26),
    ]

    COLOR_PICKERS = [
        ("상호명 색상",      "L_SHOP_NAME_COLOR"),
        ("제품명 색상",      "L_PRODUCT_COLOR"),
        ("조색번호 배경색",  "L_CODE_BG_COLOR"),
        ("조색번호 글씨색",  "L_CODE_FG_COLOR"),
        ("슬로건 색상",      "L_SLOGAN_COLOR"),
        ("날짜/시간 색상",   "L_DATETIME_COLOR"),
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.title("디자인 설정")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)
        self.temp_cfg = dict(CFG)   # 임시 작업본

        main = tk.Frame(self, bg=UI_BG)
        main.pack(padx=14, pady=14)

        # ── 좌측: 미리보기 ──
        left = tk.Frame(main, bg=UI_BG)
        left.grid(row=0, column=0, sticky='n', padx=(0, 16))
        tk.Label(left, text="미리보기", font=("맑은 고딕", 10, "bold"),
                 bg=UI_BG, fg=UI_TEXT).pack(anchor='w', pady=(0,4))
        self.preview_label = tk.Label(left, bg=UI_CARD)
        self.preview_label.pack()

        # ── 우측: 조정 패널 (스크롤 가능) ──
        right_outer = tk.Frame(main, bg=UI_BG)
        right_outer.grid(row=0, column=1, sticky='n')

        canvas = tk.Canvas(right_outer, width=380, height=480, highlightthickness=0, bg=UI_BG)
        vscroll = tk.Scrollbar(right_outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)

        right = tk.Frame(canvas, bg=UI_BG)
        canvas.create_window((0,0), window=right, anchor='nw')
        right.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        # 테두리 스타일
        tk.Label(right, text="테두리 스타일", font=("맑은 고딕", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky='w', pady=(0,2))
        self.border_var = tk.StringVar(value=self.temp_cfg.get('L_BORDER_STYLE', 'classic'))
        border_box = ttk.Combobox(right, textvariable=self.border_var, state='readonly',
                                    values=['classic', 'boxed', 'simple'], width=20)
        border_box.grid(row=1, column=0, columnspan=2, sticky='w', pady=(0,8))
        border_box.bind('<<ComboboxSelected>>', lambda e: self.update_preview())

        # 장식선 스타일
        tk.Label(right, text="구분선 장식", font=("맑은 고딕", 10, "bold")).grid(
            row=2, column=0, columnspan=2, sticky='w', pady=(0,2))
        self.deco_var = tk.StringVar(value=self.temp_cfg.get('L_DECORATION', 'wave'))
        deco_box = ttk.Combobox(right, textvariable=self.deco_var, state='readonly',
                                  values=['wave', 'line', 'none'], width=20)
        deco_box.grid(row=3, column=0, columnspan=2, sticky='w', pady=(0,10))
        deco_box.bind('<<ComboboxSelected>>', lambda e: self.update_preview())

        # 슬라이더들
        self.slider_vars = {}
        r = 4
        for label, key, lo, hi in self.SLIDERS:
            tk.Label(right, text=label, font=("맑은 고딕", 9)).grid(
                row=r, column=0, columnspan=2, sticky='w', pady=(4,0))
            r += 1
            var = tk.IntVar(value=int(self.temp_cfg.get(key, DEFAULT_CONFIG[key])))
            self.slider_vars[key] = var
            s = tk.Scale(right, from_=lo, to=hi, orient=tk.HORIZONTAL,
                         variable=var, length=300, showvalue=True,
                         command=lambda v, k=key: self.on_slider_change(k))
            s.grid(row=r, column=0, columnspan=2, sticky='w')
            r += 1

        # 색상 선택
        tk.Label(right, text="색상", font=("맑은 고딕", 10, "bold")).grid(
            row=r, column=0, columnspan=2, sticky='w', pady=(10,2))
        r += 1
        self.color_buttons = {}
        for label, key in self.COLOR_PICKERS:
            tk.Label(right, text=label, font=("맑은 고딕", 9)).grid(
                row=r, column=0, sticky='w', pady=3)
            cur = self.temp_cfg.get(key, DEFAULT_CONFIG[key])
            btn = tk.Button(right, width=8, bg=cur,
                            command=lambda k=key: self.pick_color(k))
            btn.grid(row=r, column=1, sticky='w', pady=3)
            self.color_buttons[key] = btn
            r += 1

        # 하단 버튼
        btn_frame = tk.Frame(self, bg=UI_BG)
        btn_frame.pack(pady=(8, 14))
        flat_button(btn_frame, "기본값으로 초기화", self.reset_default).pack(side=tk.LEFT, padx=6)
        flat_button(btn_frame, "저장", self.on_save, primary=True).pack(side=tk.LEFT, padx=6)
        flat_button(btn_frame, "취소", self.destroy).pack(side=tk.LEFT, padx=6)

        self.update_preview()

    def on_slider_change(self, key):
        self.temp_cfg[key] = self.slider_vars[key].get()
        self.update_preview()

    def pick_color(self, key):
        from tkinter import colorchooser
        cur = self.temp_cfg.get(key, DEFAULT_CONFIG[key])
        result = colorchooser.askcolor(color=cur, parent=self, title="색상 선택")
        if result and result[1]:
            self.temp_cfg[key] = result[1]
            self.color_buttons[key].config(bg=result[1])
            self.update_preview()

    def reset_default(self):
        for key, var in self.slider_vars.items():
            var.set(DEFAULT_CONFIG[key])
            self.temp_cfg[key] = DEFAULT_CONFIG[key]
        for key, btn in self.color_buttons.items():
            btn.config(bg=DEFAULT_CONFIG[key])
            self.temp_cfg[key] = DEFAULT_CONFIG[key]
        self.border_var.set(DEFAULT_CONFIG['L_BORDER_STYLE'])
        self.deco_var.set(DEFAULT_CONFIG['L_DECORATION'])
        self.temp_cfg['L_BORDER_STYLE'] = DEFAULT_CONFIG['L_BORDER_STYLE']
        self.temp_cfg['L_DECORATION']   = DEFAULT_CONFIG['L_DECORATION']
        self.update_preview()

    def update_preview(self):
        self.temp_cfg['L_BORDER_STYLE'] = self.border_var.get()
        self.temp_cfg['L_DECORATION']   = self.deco_var.get()

        global CFG
        backup = dict(CFG)
        CFG = self.temp_cfg
        try:
            img = draw_label(self.SAMPLE_DATA)
        finally:
            CFG = backup

        ratio = 420 / img.width
        preview = img.resize((420, int(img.height * ratio)), Image.LANCZOS)
        self.tk_preview = ImageTk.PhotoImage(preview)
        self.preview_label.config(image=self.tk_preview)

    def on_save(self):
        global CFG
        CFG.update(self.temp_cfg)
        save_config(CFG)
        messagebox.showinfo("저장 완료", "디자인 설정이 저장되었습니다.", parent=self)
        self.destroy()

# ── 설정 화면 ──────────────────────────────────────────
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("설정")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        self.configure(fg_color=UI_BG)

        fields = [
            ("조색 이력 파일 경로",  "HISTORY_PATH"),
            ("수성 폴더 경로",        "ROOT_SU"),
            ("유성 폴더 경로",        "ROOT_YU"),
            ("폰트 파일 경로",        "FONT_PATH"),
            ("상호명",               "SHOP_NAME"),
            ("슬로건",               "SHOP_SLOGAN"),
            ("주소",                 "SHOP_ADDR"),
            ("전화번호",             "SHOP_TEL"),
            ("팩스번호",             "SHOP_FAX"),
            ("고객정보 조회 비밀번호", "CUSTOMER_PIN"),
        ]

        self.vars = {}
        for i, (label, key) in enumerate(fields):
            tk.Label(self, text=label, font=("맑은 고딕", 10), bg=UI_BG, fg=UI_TEXT,
                     anchor='w').grid(row=i, column=0, padx=(16,8), pady=4, sticky='w')
            var = tk.StringVar(value=get_cfg(key))
            tk.Entry(self, textvariable=var, width=42,
                     font=("맑은 고딕", 10)).grid(row=i, column=1, padx=(0,16), pady=4)
            self.vars[key] = var

        btn_frame = tk.Frame(self, bg=UI_BG)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=12)
        flat_button(btn_frame, "저장", self.on_save, primary=True).pack(side=tk.LEFT, padx=8)
        flat_button(btn_frame, "취소", self.destroy).pack(side=tk.LEFT, padx=8)

    def on_save(self):
        global CFG
        for key, var in self.vars.items():
            CFG[key] = var.get()
        save_config(CFG)
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.", parent=self)
        self.destroy()

# ── 메인 GUI ───────────────────────────────────────────
class LabelApp:
    def __init__(self, root):
        self.root = root
        root.title("제비스코 라벨 프린터")
        root.geometry("440x560")
        root.resizable(False, False)
        root.configure(fg_color=UI_BG)
        root.attributes('-topmost', True)

        outer = ctk.CTkFrame(root, fg_color=UI_BG)
        outer.pack(fill="both", expand=True, padx=24, pady=24)

        # ── 헤더: 동그란 아이콘 + 상호명 ──
        icon_circle = ctk.CTkLabel(outer, text="P", font=("맑은 고딕", 18, "bold"),
                                    fg_color=UI_CARD, text_color=UI_TEXT,
                                    width=48, height=48, corner_radius=24)
        icon_circle.pack(pady=(4, 12))

        ctk.CTkLabel(outer, text=get_cfg('SHOP_NAME'),
                     font=("맑은 고딕", 16, "bold"),
                     fg_color=UI_BG, text_color=UI_TEXT).pack()
        ctk.CTkLabel(outer, text="조색 완료 후 버튼을 눌러주세요",
                     font=("맑은 고딕", 11),
                     fg_color=UI_BG, text_color=UI_MUTED).pack(pady=(4, 0))

        sep = ctk.CTkFrame(outer, fg_color=UI_BORDER, height=1, corner_radius=0)
        sep.pack(fill="x", pady=18)

        # ── 메뉴 버튼들 (카드형 리스트) ──
        menu_frame = ctk.CTkFrame(outer, fg_color=UI_BG)
        menu_frame.pack(fill="x")

        self._menu_button(menu_frame, "라벨 출력", self.do_print_last, primary=True)
        self._menu_button(menu_frame, "이전 목록", self.do_history)
        self._menu_button(menu_frame, "고객 등록", self.do_customer)
        self._menu_button(menu_frame, "환경 설정", self.do_settings)
        self._menu_button(menu_frame, "라벨 디자인 설정", self.do_design_settings)

        # ── 상태 표시 바 ──
        status_card = ctk.CTkFrame(outer, fg_color=UI_CARD, corner_radius=10,
                                    border_width=1, border_color=UI_BORDER)
        status_card.pack(fill="x", pady=(16, 0))
        self.status = ctk.CTkLabel(status_card, text="대기 중",
                                    font=("맑은 고딕", 11),
                                    fg_color=UI_CARD, text_color=UI_MUTED)
        self.status.pack(pady=10)

    def _menu_button(self, parent, text, command, primary=False):
        bg = UI_DARK_BTN if primary else UI_CARD
        fg = '#ffffff' if primary else UI_TEXT
        hover = UI_DARK_BTN_HOVER if primary else '#f0eee7'

        btn = ctk.CTkButton(parent, text=f"{text}          >", command=command,
                             font=("맑은 고딕", 13), fg_color=bg, hover_color=hover,
                             text_color=fg, corner_radius=10,
                             border_width=0 if primary else 1, border_color=UI_BORDER,
                             height=48, anchor='w')
        btn.pack(fill="x", pady=4)

    def do_print_last(self):
        try:
            self.status.configure(text="데이터 읽는 중...", text_color=UI_TEXT)
            self.root.update()
            records = get_records(1)
            self._show_preview(records[0])
        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.status.configure(text="오류 발생", text_color='#c0392b')

    def do_history(self):
        try:
            dlg = HistoryDialog(self.root)
            self.root.wait_window(dlg)
            if dlg.selected:
                self._show_preview(dlg.selected)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.status.configure(text="오류 발생", text_color='#c0392b')

    def do_settings(self):
        if getattr(self, '_settings_dlg', None) is not None and self._settings_dlg.winfo_exists():
            self._settings_dlg.lift()
            self._settings_dlg.focus_force()
            return
        self._settings_dlg = SettingsDialog(self.root)
        self.root.wait_window(self._settings_dlg)
        self._settings_dlg = None

    def do_design_settings(self):
        if getattr(self, '_design_dlg', None) is not None and self._design_dlg.winfo_exists():
            self._design_dlg.lift()
            self._design_dlg.focus_force()
            return
        self._design_dlg = DesignSettingsDialog(self.root)
        self.root.wait_window(self._design_dlg)
        self._design_dlg = None

    def do_customer(self):
        if getattr(self, '_customer_dlg', None) is not None and self._customer_dlg.winfo_exists():
            self._customer_dlg.lift()
            self._customer_dlg.focus_force()
            return
        self._customer_dlg = CustomerDialog(self.root)
        self.root.wait_window(self._customer_dlg)
        self._customer_dlg = None

    def _show_preview(self, rec):
        try:
            data = parse_record(rec)
            dlg  = PrintPreviewDialog(self.root, data)
            self.root.wait_window(dlg)
            if dlg.result:
                self.status.configure(
                    text=f"출력 완료  [ {data['color_code']} ]", text_color='#0f766e')
            else:
                self.status.configure(text="취소됨", text_color=UI_MUTED)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.status.configure(text="오류 발생", text_color='#c0392b')


if __name__ == '__main__':
    root = ctk.CTk()
    app = LabelApp(root)
    root.mainloop()
