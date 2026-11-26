# -*- coding: utf-8 -*-
"""
Công cụ đồng bộ tự động 1-Click:
Đọc dữ liệu từ 'Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx' và tự động cập nhật:
1. Gia_Pha_Ho_Bach.docx (File Word hoàn chỉnh)
2. Gia_Pha_Ho_Bach.md (File Markdown kèm sơ đồ Mermaid)
3. Gia_Pha_Viewer.html (Web viewer tra cứu offline)
"""
import os
import re
import json
import openpyxl
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx")
DOCX_PATH = os.path.join(BASE_DIR, "Gia_Pha_Ho_Bach.docx")
MD_PATH = os.path.join(BASE_DIR, "Gia_Pha_Ho_Bach.md")
HTML_PATH = os.path.join(BASE_DIR, "Gia_Pha_Viewer.html")

print("=" * 65)
print("  HỆ THỐNG ĐỒNG BỘ GIA PHẢ HỌ BẠCH TỪ EXCEL")
print("  Nguồn dữ liệu: Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx")
print("=" * 65)

if not os.path.exists(EXCEL_PATH):
    print(f"LỖI: Không tìm thấy file Excel tại {EXCEL_PATH}")
    input("Nhấn Enter để thoát...")
    exit(1)

# 1. ĐỌC DỮ LIỆU TỪ EXCEL
print("--> Đang đọc dữ liệu từ Excel...")
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

# Đọc Sheet Toàn bộ nhân đinh
ws_all = wb["Toàn bộ nhân đinh"]
records = []
for row_idx in range(5, ws_all.max_row + 1):
    vals = [ws_all.cell(row=row_idx, column=c).value for c in range(1, 23)]
    # Bỏ qua dòng trống nếu không có họ tên
    if not vals[4] or str(vals[4]).strip() == "":
        continue
    
    id_cha_raw = vals[15] if len(vals) > 15 and vals[15] is not None else 0
    id_cha_str = str(id_cha_raw).strip()
    if id_cha_str.isdigit():
        id_cha_val = int(id_cha_str)
    else:
        id_cha_val = id_cha_str if id_cha_str and id_cha_str != "0" else 0

    records.append({
        "stt": vals[0] if vals[0] is not None else len(records) + 1,
        "ma": str(vals[1] or "").strip(),
        "can": str(vals[2] or "").strip(),
        "doi": str(vals[3] or "").strip(),
        "ho_ten": str(vals[4] or "").strip(),
        "huy": str(vals[5] or "").strip(),
        "gioi_tinh": str(vals[6] or "").strip(),
        "nam_sinh": str(vals[7] or "").strip(),
        "nam_mat": str(vals[8] or "").strip(),
        "ngay_gio": str(vals[9] or "").strip(),
        "cha": str(vals[10] or "").strip(),
        "me": str(vals[11] or "").strip(),
        "phoi_ngau": str(vals[12] or "").strip(),
        "mo_phan": str(vals[13] or "").strip(),
        "ghi_chu": str(vals[14] or "").strip(),
        "id_cha": id_cha_val,
        "ngay_mat": str(vals[16] or "").strip() if len(vals) > 16 else "",
        "sdt": str(vals[17] or "").strip() if len(vals) > 17 else "",
        "facebook": str(vals[18] or "").strip() if len(vals) > 18 else "",
        "dia_chi": str(vals[19] or "").strip() if len(vals) > 19 else "",
        "nghe_nghiep": str(vals[20] or "").strip() if len(vals) > 20 else "",
        "thu_tu": str(vals[21] or "").strip() if len(vals) > 21 else "",
    })

# Tạo bảng tra cứu stt -> ho_ten và ma -> ho_ten để tự động điền tên cha nếu cột Cha bị trống
stt_to_name = {r["stt"]: r["ho_ten"] for r in records}
ma_to_name = {r["ma"]: r["ho_ten"] for r in records if r["ma"]}
for r in records:
    if (not r["cha"] or r["cha"].startswith("=")) and r["id_cha"]:
        if r["id_cha"] in stt_to_name:
            r["cha"] = stt_to_name[r["id_cha"]]
        elif r["id_cha"] in ma_to_name:
            r["cha"] = ma_to_name[r["id_cha"]]

print(f"    + Đã tải {len(records)} nhân đinh (kèm ID_CHA).")

# -------------------------------------------------------------
# ĐỌC VÀ TỰ ĐỘNG TỔNG HỢP LỊCH KỴ GIỖ
# -------------------------------------------------------------
def clean_person_name(name):
    if not name:
        return ""
    s = re.sub(r'^(Cụ\s+bà|Cụ|Ông|Bà|Liệt\s+sỹ)\s*', '', str(name), flags=re.IGNORECASE).strip()
    s = re.sub(r'\(.*?\)', '', s).strip()
    return ' '.join(s.lower().split())

def parse_lunar_date(date_str):
    if not date_str:
        return None, None
    s = str(date_str).strip()
    s_clean = s.replace("ÂL", "").replace("AL", "").replace("âl", "").replace("al", "").strip()
    if "/" in s_clean:
        parts = s_clean.split("/")
        try:
            d = int(parts[0].strip())
            m = int(parts[1].strip())
            return f"{d:02d}", m
        except (ValueError, IndexError):
            pass
    if "tháng" in s:
        parts = s.split("tháng")
        d_part = parts[0].replace("Ngày", "").replace("ngày", "").strip()
        m_part = parts[1].strip()
        m_map = {
            "Mười Một": 11, "Mười Hai": 12, "Mười": 10, "Giêng": 1, "Hai": 2, "Ba": 3,
            "Tư": 4, "Năm": 5, "Sáu": 6, "Bảy": 7, "Tám": 8, "Chín": 9, "Chạp": 12
        }
        m = None
        for k, v in m_map.items():
            if k in m_part:
                m = v
                break
        if m is None:
            digits = "".join([c for c in m_part if c.isdigit()])
            if digits:
                m = int(digits)
        try:
            d_digits = "".join([c for c in d_part if c.isdigit()])
            d = int(d_digits) if d_digits else 1
            return f"{d:02d}", m or 1
        except (ValueError, IndexError):
            pass
    return None, None

month_vietnamese = {
    1: "tháng Giêng", 2: "tháng Hai", 3: "tháng Ba", 4: "tháng Tư",
    5: "tháng Năm", 6: "tháng Sáu", 7: "tháng Bảy", 8: "tháng Tám",
    9: "tháng Chín", 10: "tháng Mười", 11: "tháng Mười Một", 12: "tháng Chạp"
}

# 1. Đọc từ Sheet Bảng Kỵ Giỗ nếu có
ws_gio = wb["Bảng Kỵ Giỗ"] if "Bảng Kỵ Giỗ" in wb.sheetnames else None
gio_items = []
seen_keys = set()

if ws_gio:
    for row_idx in range(5, ws_gio.max_row + 1):
        vals = [ws_gio.cell(row=row_idx, column=c).value for c in range(1, 7)]
        if not vals[2] or str(vals[2]).strip() == "":
            continue
        date_str = str(vals[1] or "").strip()
        name = str(vals[2] or "").strip()
        can = str(vals[3] or "").strip()
        mo = str(vals[4] or "").strip()
        ghi = str(vals[5] or "").strip()
        
        d_str, m_num = parse_lunar_date(date_str)
        if d_str and m_num:
            name_clean = clean_person_name(name)
            name_core = re.sub(r'\(.*?\)', '', name_clean).strip()
            seen_keys.add((d_str, m_num, name_clean))
            seen_keys.add((d_str, m_num, name_core))
            gio_items.append({
                "thang": m_num,
                "ngay": d_str,
                "ten": name,
                "can": can,
                "mo": mo,
                "ghi": ghi,
                "ngay_str": date_str
            })

# 2. Quét tự động từ Sheet Toàn bộ nhân đinh (nếu nam đinh có Ngày giỗ ÂL thì tự động đưa vào)
auto_added_count = 0
for r in records:
    if r["ngay_gio"]:
        d_str, m_num = parse_lunar_date(r["ngay_gio"])
        if d_str and m_num:
            name_clean = clean_person_name(r["ho_ten"])
            name_core = re.sub(r'\(.*?\)', '', name_clean).strip()
            if (d_str, m_num, name_clean) not in seen_keys and (d_str, m_num, name_core) not in seen_keys:
                seen_keys.add((d_str, m_num, name_clean))
                seen_keys.add((d_str, m_num, name_core))
                
                note_parts = []
                if r["cha"]: note_parts.append(f"Con ông {r['cha']}")
                if r["me"]: note_parts.append(f"Mẹ: {r['me']}")
                if r["ghi_chu"]: note_parts.append(r["ghi_chu"])
                note_text = " • ".join(note_parts)
                
                date_display = f"{d_str} {month_vietnamese.get(m_num, f'tháng {m_num}')}"
                gio_items.append({
                    "thang": m_num,
                    "ngay": d_str,
                    "ten": r["ho_ten"],
                    "can": f"{r['can']} - {r['doi']}",
                    "mo": r["mo_phan"],
                    "ghi": note_text,
                    "ngay_str": date_display
                })
                auto_added_count += 1

# 3. Quét tự động từ Cột Phối ngẫu trong Sheet Toàn bộ nhân đinh (phát hiện vợ/chồng có ngày mất ÂL)
spouse_added_count = 0
for r in records:
    pn = r.get("phoi_ngau", "")
    if pn:
        parts = pn.split(";")
        for part in parts:
            part = part.strip()
            if not part:
                continue
            lower = part.lower()
            if not any(k in lower for k in ["mất", "giỗ"]):
                continue
            date_match = re.search(r'giỗ\s*([0-9]{1,2}\s*/\s*[0-9]{1,2})\s*(?:âl|al)?', lower)
            if not date_match:
                date_match = re.search(r'giỗ\s*(?:ngày\s*)?([0-9]{1,2}\s*tháng\s*(?:[0-9]{1,2}|giêng|chạp))', lower)
            if not date_match:
                date_match = re.search(r'mất\s*([0-9]{1,2}\s*/\s*[0-9]{1,2})\s*(?:âl|al)?', lower)
            if not date_match:
                date_match = re.search(r'mất\s*(?:ngày\s*)?([0-9]{1,2}\s*tháng\s*(?:[0-9]{1,2}|giêng|chạp))', lower)
            if not date_match:
                continue
            date_raw = date_match.group(1).replace(" ", "")
            d_str, m_num = parse_lunar_date(date_raw)
            if not (d_str and m_num):
                continue
            
            clean_part = re.sub(r'^(Bà\s+\d+:|Bà\s+cả:|Bà\s+hai:|Bà\s+ba:|Bà\s+tư:)\s*', '', part).strip()
            name = clean_part.split('(')[0].split(',')[0].strip()
            name_clean = clean_person_name(name)
            name_core = re.sub(r'\(.*?\)', '', name_clean).strip()
            
            if (d_str, m_num, name_clean) not in seen_keys and (d_str, m_num, name_core) not in seen_keys:
                seen_keys.add((d_str, m_num, name_clean))
                seen_keys.add((d_str, m_num, name_core))
                
                title = f"Cụ bà {name}" if any(f"Đời {i}" in r['doi'] for i in range(1, 7)) else f"Bà {name}"
                can_doi = f"{r['can']} - {r['doi']}".strip(" -")
                note_text = f"{part}. Vợ ông {r['ho_ten']}"
                date_display = f"{d_str} {month_vietnamese.get(m_num, f'tháng {m_num}')}"
                
                gio_items.append({
                    "thang": m_num,
                    "ngay": d_str,
                    "ten": title,
                    "can": can_doi,
                    "mo": "",
                    "ghi": note_text,
                    "ngay_str": date_display
                })
                spouse_added_count += 1
                auto_added_count += 1

# Sắp xếp tuần tự theo 12 tháng Âm lịch (từ tháng 1 đến tháng 12, ngày từ 1 đến 31)
gio_items.sort(key=lambda x: (x["thang"], int(x["ngay"])))

gio_list = []
for idx, it in enumerate(gio_items, start=1):
    gio_list.append((
        idx,
        it["ngay_str"],
        it["ten"],
        it["can"],
        it["mo"],
        it["ghi"]
    ))

print(f"    + Đã tải và tự động tổng hợp {len(gio_list)} ngày kỵ giỗ (tự động phát hiện thêm {auto_added_count} người từ danh bạ).")

# -------------------------------------------------------------
# 2. XUẤT BẢN FILE WORD (.DOCX)
# -------------------------------------------------------------
print("--> Đang tái tạo file Word (Gia_Pha_Ho_Bach.docx)...")
doc = Document()
for section in doc.sections:
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

COLOR_PRIMARY = RGBColor(31, 73, 125)
COLOR_SECONDARY = RGBColor(89, 89, 89)
COLOR_TEXT = RGBColor(38, 38, 38)

def add_h1(text):
    h = doc.add_heading(level=1)
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

def add_h2(text):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    r = h.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

def add_p(text, bold_prefix=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(11.5)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_TEXT
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11.5)
    r.font.italic = italic
    r.font.color.rgb = COLOR_TEXT
    return p

def add_callout(text, title=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(r'<w:shd {} w:fill="F2F5F8"/>'.format(nsdecls('w')))
    tcPr.append(shd)
    borders = parse_xml(r'''
        <w:tcBorders {} >
            <w:left w:val="single" w:sz="24" w:space="0" w:color="1F497D"/>
            <w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/>
        </w:tcBorders>
    '''.format(nsdecls('w')))
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.15)
    if title:
        rt = p.add_run(title + "\n")
        rt.font.name = "Times New Roman"
        rt.font.size = Pt(12)
        rt.font.bold = True
        rt.font.color.rgb = COLOR_PRIMARY
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.font.italic = True
    r.font.color.rgb = COLOR_TEXT

# Trang bìa
for _ in range(3): doc.add_paragraph()
p_cov = doc.add_paragraph()
p_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_cov.add_run("BẠCH TỘC ĐẠI TÔN\n")
r.font.name = "Times New Roman"; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = COLOR_SECONDARY
r2 = p_cov.add_run("TÔN ĐỒ TỰ PHẢ\nTIỂU CHI 1 - TRUNG CHI 2\n")
r2.font.name = "Times New Roman"; r2.font.size = Pt(26); r2.font.bold = True; r2.font.color.rgb = COLOR_PRIMARY
r3 = p_cov.add_run(f"(DỮ LIỆU ĐỒNG BỘ TỰ ĐỘNG - TỔNG SỐ {len(records)} NHÂN ĐINH)")
r3.font.name = "Times New Roman"; r3.font.size = Pt(12); r3.font.italic = True; r3.font.color.rgb = COLOR_SECONDARY

for _ in range(4): doc.add_paragraph()
p_loc = doc.add_paragraph()
p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_loc.add_run("LÀNG PHÚ ĐIỀN - XÃ HƯNG PHÚ - HUYỆN HƯNG NGUYÊN - TỈNH NGHỆ AN\nHIỆU ĐÍNH & CẬP NHẬT TỰ ĐỘNG NĂM 2026")
r.font.name = "Times New Roman"; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = COLOR_TEXT

doc.add_page_break()

# Lời nói đầu & Lịch sử họ Bạch
add_h1("LỜI NÓI ĐẦU")
add_callout(
    "\"Con người ta sinh ra ở đời ai cũng có Tổ, có Tông, như cây có cội, như sông có nguồn. "
    "Đó là đạo lý truyền thống ngàn đời tốt đẹp của dân tộc ta.\"",
    "Cội Nguồn Dòng Tộc"
)
add_p("Cuốn phả hệ này là Tôn Đồ Tự Phả của Tiểu Chi 1, Trung Chi 2 thuộc Bạch Tộc Đại Tôn, nguyên quán tại làng Phú Điền, xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An.")
add_p("Vào những năm 1990, ông Bạch Hưng Thủng đã có công đầu sưu tầm từ các bản Hán văn cổ của dòng tộc, dịch thuật sang chữ Quốc ngữ. Đến năm 2014 (Giáp Ngọ), ban tộc biểu gồm các ông: Bạch Hưng Thái, Bạch Hưng Phú, Bạch Hưng Văn, Bạch Hưng Uyển, Bạch Hưng Tùng đã cùng nhau đối chiếu, chỉnh lý và in ấn thành tập sách gia phả hoàn chỉnh.")
add_p("ĐẶC BIỆT LƯU Ý: Toàn bộ phả hệ được phân tích chính xác 100% dựa vào văn bản mô tả từng trang của dòng họ. Nguồn dữ liệu hiện được đồng bộ tự động từ file Excel quản trị trung tâm.")

add_h1("VÀI NÉT VỀ LỊCH SỬ HỌ BẠCH & DANH NHÂN BẠCH LIÊU")
add_h2("1. Thủy tổ Trạng nguyên Bạch Liêu (1236 - 1315)")
add_p("Dòng họ Bạch Việt Nam tự hào có Cụ Thủy tổ là Trạng nguyên Bạch Liêu, sinh ngày 15 tháng Giêng năm Bính Thìn (1236) tại làng Yên Xá, huyện Chân Phúc (nay thuộc xã Mã Thành, huyện Yên Thành, tỉnh Nghệ An). Năm Bính Dần (1266) đời vua Trần Thánh Tông, ngài đỗ Trạng nguyên (Đệ nhất giáp tiến sĩ đệ nhất danh) khi tròn 31 tuổi.")
add_p("Ngài là Môn khách danh dự của Thượng phụ Chiêu Minh Đại vương Trần Quang Khải, có công lao to lớn hoạch định cơ mưu đánh giặc Nguyên - Mông và soạn thảo 'Biến pháp tam chương' cải cách quốc gia. Cụ tạ thế ngày 24 tháng Giêng năm Ất Mão (1315), thọ 80 tuổi.")
add_h2("2. Đền thờ họ Bạch tại Hưng Phú, Hưng Nguyên")
add_p("Đền thờ họ Bạch tại xã Hưng Phú đã được Nhà nước xếp hạng Di tích Lịch sử - Văn hóa cấp Tỉnh theo Quyết định số 2015/QĐ-UBND ngày 16/12/1993, được tôn tạo khang trang năm 2013.")

# Xuất danh sách nhân đinh theo từng Can
sections_def = [
    ("TỔ KHỞI PHÁT TRUNG CHI 2 & ĐỜI THỨ NHẤT, THỨ HAI", ["Khởi phát", "Tiểu Chi 1", "Trung Chi 2"]),
    ("CAN THỨ NHẤT: TỔ CAN BẠCH HƯNG CƯ (ĐỜI 3 ĐẾN CÁC ĐỜI TIẾP THEO)", ["Can 1"]),
    ("CAN THỨ HAI: TỔ CAN BẠCH HƯNG KHÂU (ĐỜI 3 ĐẾN CÁC ĐỜI TIẾP THEO)", ["Can 2"]),
    ("CAN THỨ BA: TỔ CAN BẠCH HƯNG NGHI (ĐỜI 3 ĐẾN CÁC ĐỜI TIẾP THEO)", ["Can 3"]),
    ("CAN THỨ TƯ: TỔ CAN BẠCH HƯNG KẾ", ["Can 4"]),
]

for sec_title, can_keys in sections_def:
    add_h1(sec_title)
    sec_records = [r for r in records if r["can"] in can_keys]
    
    # Gom theo đời
    dois = []
    for r in sec_records:
        if r["doi"] not in dois:
            dois.append(r["doi"])
            
    for d in dois:
        d_records = [r for r in sec_records if r["doi"] == d]
        add_h2(f"Thế Hệ: {d} ({len(d_records)} người)")
        for p in d_records:
            bio_p = doc.add_paragraph()
            bio_p.paragraph_format.space_before = Pt(4)
            bio_p.paragraph_format.space_after = Pt(2)
            bio_p.paragraph_format.left_indent = Inches(0.15)
            
            # Họ tên & chức vị / húy
            run_name = bio_p.add_run(f"• {p['ho_ten']}")
            run_name.font.name = "Times New Roman"
            run_name.font.size = Pt(11)
            run_name.font.bold = True
            run_name.font.color.rgb = COLOR_PRIMARY
            
            meta_str = ""
            if p["huy"]: meta_str += f" (Húy: {p['huy']})"
            if p["nam_sinh"] or p["nam_mat"]:
                meta_str += f" [{p['nam_sinh'] or '?'} - {p['nam_mat'] or '?'}]"
            if p["ngay_gio"]: meta_str += f" - Giỗ: {p['ngay_gio']}"
            if p["cha"]: meta_str += f" - Con ông: {p['cha']}"
            if p["phoi_ngau"]: meta_str += f" - Phối ngẫu: {p['phoi_ngau']}"
            if p["mo_phan"]: meta_str += f" - Mộ: {p['mo_phan']}"
            
            run_meta = bio_p.add_run(meta_str + "\n")
            run_meta.font.name = "Times New Roman"
            run_meta.font.size = Pt(10.5)
            run_meta.font.italic = True
            
            if p["ghi_chu"]:
                run_note = bio_p.add_run(f"  Ghi chú: {p['ghi_chu']}")
                run_note.font.name = "Times New Roman"
                run_note.font.size = Pt(10.5)
                run_note.font.color.rgb = COLOR_TEXT

doc.save(DOCX_PATH)
print("    + Đã lưu Gia_Pha_Ho_Bach.docx thành công.")

# -------------------------------------------------------------
# 3. XUẤT BẢN FILE MARKDOWN (.MD)
# -------------------------------------------------------------
print("--> Đang tái tạo file Markdown (Gia_Pha_Ho_Bach.md)...")
md_lines = [
    "# GIA PHẢ HỌ BẠCH",
    "## TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2 - THUỘC BẠCH TỘC ĐẠI TÔN",
    "**Địa danh:** Làng Phú Điền, Xã Hưng Phú, Huyện Hưng Nguyên, Tỉnh Nghệ An  ",
    f"**Thời gian biên soạn:** Năm Giáp Ngọ (2014) - Cập nhật tự động: Năm 2026 (Tổng số: {len(records)} nhân đinh)  ",
    "**Nguyên tắc số hóa:** **Phân tích quan hệ huyết thống 100% DỰA VÀO VĂN BẢN (TEXT), LOẠI BỎ SAI LỆCH DO SƠ ĐỒ VẼ GỘP CỘT NĂM 2014.**\n",
    "---\n",
    "> ### ⚠️ LƯU Ý ĐẶC BIỆT VỀ QUY TẮC PHÂN TÍCH QUAN HỆ HUYẾT THỐNG",
    "> Trong bản in sách gia phả năm 2014, các trang sơ đồ vẽ hình khối do người vẽ dồn cột giản lược nên đã xảy ra sự sai lệch nghiêm trọng về quan hệ cha con giữa các nhánh.",
    "> Toàn bộ phả hệ ở đây đã được hiệu chỉnh chính xác 100% theo lời văn mô tả của từng trang sách và được đồng bộ từ file Excel quản trị trung tâm.\n",
    "---\n",
    "## SƠ ĐỒ PHẢ HỆ TỔNG THỂ",
    "```mermaid",
    "graph TD",
    '    BHH["TỔ KHỞI PHÁT TRUNG CHI 2<br/><b>BẠCH HƯNG HÓA</b> (1730 - 1802)<br/>Giỗ: 14/4 AL | Mộ: Cồn Xanh"]',
    '    BTT["ĐỜI 1 - TIỂU CHI 1<br/><b>BẠCH TRẦN TẠO</b> (1772 - 1801)<br/>Giỗ: 28/9 AL | Vợ: Phạm Thị Hàng Nhất"]',
    '    BHH --> BTT',
    '    BHKham["ĐỜI 2<br/>Bạch Hưng Khâm<br/>(Mất năm 20 tuổi)"]',
    '    BHLuu["ĐỜI 2<br/><b>BẠCH HƯNG LƯU</b> (1798 - 1873)<br/>Húy Túc | Thọ 76 | Giỗ: 29/2 AL<br/>Vợ: Phạm Thị Biềng"]',
    '    BTT --> BHKham',
    '    BTT --> BHLuu',
    '    C1["CAN 1<br/><b>BẠCH HƯNG CƯ</b><br/>(Húy Khuơ, 1819)<br/>Giỗ: 9/10 AL"]',
    '    C2["CAN 2<br/><b>BẠCH HƯNG KHÂU</b><br/>(Húy Hoài, 1821-1897)<br/>Giỗ: 2/8 AL"]',
    '    C3["CAN 3<br/><b>BẠCH HƯNG NGHI</b><br/>(1829, thọ 89)<br/>Giỗ: 10/4 AL"]',
    '    C4["CAN 4<br/><b>BẠCH HƯNG KẾ</b><br/>(Húy Cơ, 1854)<br/>Vô tự - Phú tự"]',
    '    BHLuu --> C1',
    '    BHLuu --> C2',
    '    BHLuu --> C3',
    '    BHLuu --> C4',
    "```\n",
    "---"
]

for sec_title, can_keys in sections_def:
    md_lines.append(f"\n## {sec_title}\n")
    sec_records = [r for r in records if r["can"] in can_keys]
    
    dois = []
    for r in sec_records:
        if r["doi"] not in dois:
            dois.append(r["doi"])
            
    for d in dois:
        d_records = [r for r in sec_records if r["doi"] == d]
        md_lines.append(f"### Thế Hệ: {d}\n")
        for p in d_records:
            name_str = f"**{p['ho_ten']}**"
            if p["huy"]: name_str += f" (Húy / Biệt danh: *{p['huy']}*)"
            md_lines.append(f"- {name_str}:")
            if p["nam_sinh"] or p["nam_mat"]:
                md_lines.append(f"  - Năm sinh/mất: {p['nam_sinh'] or '?'} - {p['nam_mat'] or '?'}")
            if p["ngay_gio"]:
                md_lines.append(f"  - Ngày giỗ: **{p['ngay_gio']}**")
            if p["cha"]:
                md_lines.append(f"  - Thân phụ (Cha): {p['cha']}" + (f" | Thân mẫu (Mẹ): {p['me']}" if p["me"] else ""))
            if p["phoi_ngau"]:
                md_lines.append(f"  - Phối ngẫu: {p['phoi_ngau']}")
            if p["mo_phan"]:
                md_lines.append(f"  - Mộ phần: {p['mo_phan']}")
            if p["ghi_chu"]:
                md_lines.append(f"  - Ghi chú: {p['ghi_chu']}")
            md_lines.append("")

if gio_list:
    md_lines.append("\n## BẢNG TRA CỨU NGÀY KỴ GIỖ ÂM LỊCH TOÀN DÒNG HỌ\n")
    md_lines.append("| STT | Ngày Âm lịch | Người được giỗ | Đời / Can | Nơi an táng / Mộ phần | Ghi chú kỵ giỗ |")
    md_lines.append("| :---: | :--- | :--- | :---: | :--- | :--- |")
    for it in gio_list:
        md_lines.append(f"| {it[0]} | **{it[1]}** | {it[2]} | {it[3]} | {it[4]} | {it[5]} |")

with open(MD_PATH, "w", encoding="utf-8") as f:
    f.write("\n".join(md_lines))
print("    + Đã lưu Gia_Pha_Ho_Bach.md thành công.")

# -------------------------------------------------------------
# 4. XUẤT BẢN FILE WEB VIEWER (.HTML)
# -------------------------------------------------------------
print("--> Đang tái tạo file Web Viewer (Gia_Pha_Viewer.html)...")
js_records = []
for r in records:
    js_records.append({
        "stt": r["stt"],
        "ma": r["ma"],
        "can": r["can"],
        "doi": r["doi"],
        "ten": r["ho_ten"],
        "huy": r["huy"],
        "gioi_tinh": r["gioi_tinh"],
        "nam_sinh": r["nam_sinh"],
        "nam_mat": r["nam_mat"],
        "ngay_gio": r["ngay_gio"],
        "cha": r["cha"],
        "me": r["me"],
        "phoi_ngau": r["phoi_ngau"],
        "mo_phan": r["mo_phan"],
        "ghi_chu": r["ghi_chu"],
        "id_cha": r.get("id_cha", 0),
        "ngay_mat": r.get("ngay_mat", ""),
        "sdt": r.get("sdt", ""),
        "facebook": r.get("facebook", ""),
        "dia_chi": r.get("dia_chi", ""),
        "nghe_nghiep": r.get("nghe_nghiep", ""),
        "thu_tu": r.get("thu_tu", ""),
    })

death_person_map = {}
for r in records:
    if r["ngay_gio"]:
        d_r, m_r = parse_lunar_date(r["ngay_gio"])
        if d_r and m_r:
            death_person_map[(d_r, m_r, clean_person_name(r["ho_ten"]))] = r

js_deaths = []
for it in gio_items:
    mo_text = f" • Mộ: {it['mo']}" if (it['mo'] and it['mo'] not in it['ghi']) else ""
    matched = death_person_map.get((it["ngay"], it["thang"], clean_person_name(it["ten"])))
    js_deaths.append({
        "thang": it["thang"],
        "ngay": it["ngay"],
        "ten": it["ten"],
        "can": it["can"],
        "ngay_gio": (matched.get("ngay_gio") if matched else "") or f"{it['ngay']}/{int(it['thang']):02d} ÂL",
        "ngay_mat": matched.get("ngay_mat", "") if matched else "",
        "nam_mat": matched.get("nam_mat", "") if matched else "",
        "ghi": f"{it['ghi']}{mo_text}".strip(" •")
    })

records_json = json.dumps(js_records, ensure_ascii=False, indent=2)
deaths_json = json.dumps(js_deaths, ensure_ascii=False, indent=2)

# Đọc template HTML hiện tại và thay thế biến people và deathDays
if os.path.exists(HTML_PATH):
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html_content = f.read()
    
    # Tìm và thay thế khối const people = ... và const deathDays = ...
    import re
    html_content = re.sub(r'const people = \[[\s\S]*?\];', lambda m: f'const people = {records_json};', html_content)
    html_content = re.sub(r'const deathDays = \[[\s\S]*?\];', lambda m: f'const deathDays = {deaths_json};', html_content)
    
    # Cập nhật hiển thị số lượng ngày giỗ và nhân đinh trên HTML
    html_content = re.sub(r'Tổng cộng \d+ Ngày Giỗ', f'Tổng cộng {len(js_deaths)} Ngày Giỗ', html_content)
    html_content = re.sub(r'Lịch Kỵ Giỗ Âm Lịch \(\d+\)', f'Lịch Kỵ Giỗ Âm Lịch ({len(js_deaths)})', html_content)
    html_content = re.sub(r'Tra Cứu Nhân Đinh \(\d+\)', f'Tra Cứu Nhân Đinh ({len(js_records)})', html_content)
    html_content = re.sub(r'(<div class="stat-num">)\d+(</div>\s*<div class="stat-label">Tổng số [^<]*</div>)',
                          rf'\g<1>{len(js_records)}\g<2>', html_content)
    
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"    + Đã lưu Gia_Pha_Viewer.html thành công ({len(js_deaths)} ngày giỗ).")

print("\n" + "=" * 65)
print("  HOÀN THÀNH ĐỒNG BỘ 100% CẢ 3 TỆP:")
print("  1. Word:       Gia_Pha_Ho_Bach.docx")
print("  2. Markdown:   Gia_Pha_Ho_Bach.md")
print("  3. Web Viewer: Gia_Pha_Viewer.html")
print("=" * 65)

