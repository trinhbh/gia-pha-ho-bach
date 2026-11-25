# -*- coding: utf-8 -*-
"""
Service layer xử lý đọc/ghi Excel, sao lưu tự động và xuất bản HTML Public
cho Hệ thống Quản trị Gia Phả Họ Bạch.
"""
import os
import re
import json
import shutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx")
HTML_TEMPLATE_PATH = os.path.join(BASE_DIR, "Gia_Pha_Viewer.html")
PUBLIC_HTML_PATH = os.path.join(BASE_DIR, "Gia_Pha_Public.html")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")

MONTH_VIETNAMESE = {
    1: "tháng Giêng", 2: "tháng Hai", 3: "tháng Ba", 4: "tháng Tư",
    5: "tháng Năm", 6: "tháng Sáu", 7: "tháng Bảy", 8: "tháng Tám",
    9: "tháng Chín", 10: "tháng Mười", 11: "tháng Mười Một", 12: "tháng Chạp"
}

SUB_SHEET_MAP = {
    "Can 1": "Can 1 - Tổ Cư",
    "Can 2": "Can 2 - Tổ Khâu",
    "Can 3": "Can 3 & Can 4",
    "Can 4": "Can 3 & Can 4",
}


def clean_person_name(name):
    """Chuẩn hóa tên nhân đinh để đối chiếu trùng lặp không phân biệt hoa/thường."""
    if not name:
        return ""
    s = re.sub(r'^(Cụ\s+bà|Cụ|Ông|Bà|Liệt\s+sỹ)\s*', '', str(name), flags=re.IGNORECASE).strip()
    s = re.sub(r'\(.*?\)', '', s).strip()
    return " ".join(s.lower().split())


def parse_lunar_date(date_str):
    """Phân tích ngày tháng Âm lịch (VD: '14/04 ÂL' hoặc '14 tháng Tư') -> ('14', 4)."""
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
    if "tháng" in s.lower():
        parts = re.split(r"tháng", s, flags=re.IGNORECASE)
        d_part = re.sub(r"ngày", "", parts[0], flags=re.IGNORECASE).strip()
        m_part = parts[1].strip() if len(parts) > 1 else ""
        m_map = {
            "Mười Một": 11, "Mười Hai": 12, "Mười": 10, "Giêng": 1, "Hai": 2, "Ba": 3,
            "Tư": 4, "Năm": 5, "Sáu": 6, "Bảy": 7, "Tám": 8, "Chín": 9, "Chạp": 12
        }
        m = None
        for k, v in m_map.items():
            if k.lower() in m_part.lower():
                m = v
                break
        if m is None:
            digits = "".join(c for c in m_part if c.isdigit())
            if digits:
                m = int(digits)
        try:
            d_digits = "".join(c for c in d_part if c.isdigit())
            d = int(d_digits) if d_digits else 1
            return f"{d:02d}", m or 1
        except (ValueError, IndexError):
            pass
    return None, None


def create_backup(excel_path=EXCEL_PATH):
    """Sao lưu file Excel trước khi ghi dữ liệu mới, giữ tối đa 30 bản gần nhất."""
    if not os.path.exists(excel_path):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"Danh_Sach_Nhan_Dinh_Ho_Bach_{ts}.xlsx")
    shutil.copy2(excel_path, backup_path)

    backups = sorted(
        [os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".xlsx")],
        key=os.path.getmtime,
    )
    while len(backups) > 30:
        oldest = backups.pop(0)
        try:
            os.remove(oldest)
        except OSError:
            pass
    return backup_path


def _parse_id_cha(raw_val):
    if raw_val is None:
        return 0
    s = str(raw_val).strip()
    if s.isdigit():
        return int(s)
    return s if s and s != "0" else 0


def load_all_data(excel_path=EXCEL_PATH):
    """Đọc toàn bộ nhân đinh (21 cột) và lịch kỵ giỗ từ Excel."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel tại {excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws_all = wb["Toàn bộ nhân đinh"]

    records = []
    for row_idx in range(5, ws_all.max_row + 1):
        vals = [ws_all.cell(row=row_idx, column=c).value for c in range(1, 22)]
        if not vals[4] or str(vals[4]).strip() == "":
            continue

        stt_val = vals[0]
        try:
            stt_int = int(stt_val) if stt_val is not None else len(records) + 1
        except (ValueError, TypeError):
            stt_int = len(records) + 1

        records.append({
            "row_idx": row_idx,
            "stt": stt_int,
            "ma": str(vals[1] or "").strip(),
            "can": str(vals[2] or "").strip(),
            "doi": str(vals[3] or "").strip(),
            "ho_ten": str(vals[4] or "").strip(),
            "huy": str(vals[5] or "").strip(),
            "gioi_tinh": str(vals[6] or "Nam").strip(),
            "nam_sinh": str(vals[7] or "").strip(),
            "nam_mat": str(vals[8] or "").strip(),
            "ngay_gio": str(vals[9] or "").strip(),
            "cha": str(vals[10] or "").strip(),
            "me": str(vals[11] or "").strip(),
            "phoi_ngau": str(vals[12] or "").strip(),
            "mo_phan": str(vals[13] or "").strip(),
            "ghi_chu": str(vals[14] or "").strip(),
            "id_cha": _parse_id_cha(vals[15]),
            "ngay_mat": str(vals[16] or "").strip(),
            "sdt": str(vals[17] or "").strip(),
            "facebook": str(vals[18] or "").strip(),
            "dia_chi": str(vals[19] or "").strip(),
            "nghe_nghiep": str(vals[20] or "").strip(),
        })

    stt_to_name = {r["stt"]: r["ho_ten"] for r in records}
    ma_to_name = {r["ma"]: r["ho_ten"] for r in records if r["ma"]}
    for r in records:
        if (not r["cha"] or r["cha"].startswith("=")) and r["id_cha"]:
            if r["id_cha"] in stt_to_name:
                r["cha"] = stt_to_name[r["id_cha"]]
            elif r["id_cha"] in ma_to_name:
                r["cha"] = ma_to_name[r["id_cha"]]

    # Đọc Bảng Kỵ Giỗ
    ws_gio = wb["Bảng Kỵ Giỗ"] if "Bảng Kỵ Giỗ" in wb.sheetnames else None
    raw_deaths = []
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

            raw_deaths.append({
                "row_idx": row_idx,
                "stt": vals[0] or (len(raw_deaths) + 1),
                "ngay_str": date_str,
                "ten": name,
                "can": can,
                "mo": mo,
                "ghi": ghi,
            })

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
                    "ngay_str": date_str,
                })

    # Tự động quét thêm từ Toàn bộ nhân đinh
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
                    if r["cha"]:
                        note_parts.append(f"Con ông {r['cha']}")
                    if r["me"]:
                        note_parts.append(f"Mẹ: {r['me']}")
                    if r["ghi_chu"]:
                        note_parts.append(r["ghi_chu"])
                    date_display = f"{d_str} {MONTH_VIETNAMESE.get(m_num, f'tháng {m_num}')}"
                    gio_items.append({
                        "thang": m_num,
                        "ngay": d_str,
                        "ten": r["ho_ten"],
                        "can": f"{r['can']} - {r['doi']}",
                        "mo": r["mo_phan"],
                        "ghi": " • ".join(note_parts),
                        "ngay_str": date_display,
                    })

    # Tự động quét phối ngẫu
    for r in records:
        pn = r.get("phoi_ngau", "")
        if not pn:
            continue
        for part in pn.split(";"):
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

            clean_part = re.sub(r'^(Bà\s+\d+:|Bà\s+cả:|Bà\s+hai:|Bà\s+ba:|Bà\s+tư:)\s*', '', part, flags=re.IGNORECASE).strip()
            name = clean_part.split("(")[0].split(",")[0].strip()
            name_clean = clean_person_name(name)
            name_core = re.sub(r'\(.*?\)', '', name_clean).strip()

            if (d_str, m_num, name_clean) not in seen_keys and (d_str, m_num, name_core) not in seen_keys:
                seen_keys.add((d_str, m_num, name_clean))
                seen_keys.add((d_str, m_num, name_core))
                title = f"Cụ bà {name}" if any(f"Đời {i}" in r["doi"] for i in range(1, 7)) else f"Bà {name}"
                can_doi = f"{r['can']} - {r['doi']}".strip(" -")
                date_display = f"{d_str} {MONTH_VIETNAMESE.get(m_num, f'tháng {m_num}')}"
                gio_items.append({
                    "thang": m_num,
                    "ngay": d_str,
                    "ten": title,
                    "can": can_doi,
                    "mo": "",
                    "ghi": f"{part}. Vợ ông {r['ho_ten']}",
                    "ngay_str": date_display,
                })

    gio_items.sort(key=lambda x: (x["thang"], int(x["ngay"])))
    wb.close()
    return {"people": records, "deaths": gio_items, "raw_deaths": raw_deaths}


def _generate_ma_dd(can, doi, stt):
    can_code = "KP"
    if can == "Can 1":
        can_code = "C1"
    elif can == "Can 2":
        can_code = "C2"
    elif can == "Can 3":
        can_code = "C3"
    elif can == "Can 4":
        can_code = "C4"
    elif "Tiểu Chi" in can:
        can_code = "TC1"
    doi_digits = "".join(c for c in str(doi) if c.isdigit())
    doi_code = f"D{doi_digits}" if doi_digits else "D0"
    return f"{can_code}-{doi_code}-{stt:02d}"


def _write_person_row(ws, row_idx, vals):
    font_cell = Font(name="Times New Roman", size=11)
    align_left = Alignment(wrap_text=True, vertical="center", horizontal="left")
    align_center = Alignment(vertical="center", horizontal="center")
    for col_idx, val in enumerate(vals, start=1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.value = val
        cell.font = font_cell
        cell.alignment = align_center if col_idx in (1, 2, 4, 7, 16) else align_left


def _find_last_data_row(ws, name_col=5):
    last_row = 4
    for r in range(5, ws.max_row + 1):
        if ws.cell(row=r, column=name_col).value:
            last_row = r
    return last_row


def add_person(payload, excel_path=EXCEL_PATH):
    """Thêm mới một nhân đinh vào Excel và tự động sao lưu trước khi ghi."""
    ho_ten = str(payload.get("ho_ten") or "").strip()
    if not ho_ten:
        raise ValueError("Họ và tên không được để trống.")

    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Toàn bộ nhân đinh"]

    last_row = _find_last_data_row(ws, name_col=5)
    max_stt = 0
    stt_to_name = {}
    for r in range(5, last_row + 1):
        s = ws.cell(row=r, column=1).value
        n = ws.cell(row=r, column=5).value
        if isinstance(s, int):
            max_stt = max(max_stt, s)
            if n:
                stt_to_name[s] = str(n).strip()

    new_stt = max_stt + 1
    can = str(payload.get("can") or "Can 2").strip()
    doi = str(payload.get("doi") or "Đời 8").strip()
    ma = str(payload.get("ma") or "").strip() or _generate_ma_dd(can, doi, new_stt)
    id_cha = _parse_id_cha(payload.get("id_cha"))
    cha_name = str(payload.get("cha") or "").strip()
    if not cha_name and isinstance(id_cha, int) and id_cha in stt_to_name:
        cha_name = stt_to_name[id_cha]

    row_vals = [
        new_stt,
        ma,
        can,
        doi,
        ho_ten,
        str(payload.get("huy") or "").strip(),
        str(payload.get("gioi_tinh") or "Nam").strip(),
        str(payload.get("nam_sinh") or "").strip(),
        str(payload.get("nam_mat") or "").strip(),
        str(payload.get("ngay_gio") or "").strip(),
        cha_name,
        str(payload.get("me") or "").strip(),
        str(payload.get("phoi_ngau") or "").strip(),
        str(payload.get("mo_phan") or "").strip(),
        str(payload.get("ghi_chu") or "").strip(),
        id_cha if id_cha else 0,
        str(payload.get("ngay_mat") or "").strip(),
        str(payload.get("sdt") or "").strip(),
        str(payload.get("facebook") or "").strip(),
        str(payload.get("dia_chi") or "").strip(),
        str(payload.get("nghe_nghiep") or "").strip(),
    ]

    target_row = last_row + 1
    _write_person_row(ws, target_row, row_vals)

    # Đồng bộ vào Sheet nhánh nếu có
    sub_name = SUB_SHEET_MAP.get(can)
    if sub_name and sub_name in wb.sheetnames:
        ws_sub = wb[sub_name]
        sub_last = _find_last_data_row(ws_sub, name_col=5)
        sub_vals = list(row_vals[:16])
        sub_vals[0] = (sub_last - 4) + 1
        _write_person_row(ws_sub, sub_last + 1, sub_vals)

    wb.save(excel_path)
    wb.close()
    return new_stt


def update_person(stt, payload, excel_path=EXCEL_PATH):
    """Cập nhật thông tin nhân đinh theo STT trong Excel."""
    ho_ten = str(payload.get("ho_ten") or "").strip()
    if not ho_ten:
        raise ValueError("Họ và tên không được để trống.")

    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Toàn bộ nhân đinh"]

    target_row = None
    old_ma = ""
    stt_to_name = {}
    for r in range(5, ws.max_row + 1):
        s = ws.cell(row=r, column=1).value
        n = ws.cell(row=r, column=5).value
        if s is not None and str(s).isdigit():
            stt_to_name[int(s)] = str(n or "").strip()
            if int(s) == int(stt):
                target_row = r
                old_ma = str(ws.cell(row=r, column=2).value or "").strip()

    if target_row is None:
        wb.close()
        raise ValueError(f"Không tìm thấy nhân đinh có STT {stt}.")

    can = str(payload.get("can") or ws.cell(target_row, 3).value or "").strip()
    doi = str(payload.get("doi") or ws.cell(target_row, 4).value or "").strip()
    ma = str(payload.get("ma") or old_ma or _generate_ma_dd(can, doi, stt)).strip()
    id_cha = _parse_id_cha(payload.get("id_cha"))
    cha_name = str(payload.get("cha") or "").strip()
    if not cha_name and isinstance(id_cha, int) and id_cha in stt_to_name:
        cha_name = stt_to_name[id_cha]

    row_vals = [
        int(stt),
        ma,
        can,
        doi,
        ho_ten,
        str(payload.get("huy") or "").strip(),
        str(payload.get("gioi_tinh") or "Nam").strip(),
        str(payload.get("nam_sinh") or "").strip(),
        str(payload.get("nam_mat") or "").strip(),
        str(payload.get("ngay_gio") or "").strip(),
        cha_name,
        str(payload.get("me") or "").strip(),
        str(payload.get("phoi_ngau") or "").strip(),
        str(payload.get("mo_phan") or "").strip(),
        str(payload.get("ghi_chu") or "").strip(),
        id_cha if id_cha else 0,
        str(payload.get("ngay_mat") or "").strip(),
        str(payload.get("sdt") or "").strip(),
        str(payload.get("facebook") or "").strip(),
        str(payload.get("dia_chi") or "").strip(),
        str(payload.get("nghe_nghiep") or "").strip(),
    ]
    _write_person_row(ws, target_row, row_vals)

    # Cập nhật tên Cha cho các con nếu tên nhân đinh thay đổi
    for r in range(5, ws.max_row + 1):
        child_id_cha = _parse_id_cha(ws.cell(row=r, column=16).value)
        if child_id_cha == int(stt):
            ws.cell(row=r, column=11).value = ho_ten

    # Đồng bộ vào sub-sheet nếu khớp Mã ĐD
    if old_ma:
        for sname in ("Can 1 - Tổ Cư", "Can 2 - Tổ Khâu", "Can 3 & Can 4"):
            if sname in wb.sheetnames:
                ws_sub = wb[sname]
                for r in range(5, ws_sub.max_row + 1):
                    if str(ws_sub.cell(row=r, column=2).value or "").strip() == old_ma:
                        sub_vals = list(row_vals[:16])
                        sub_vals[0] = ws_sub.cell(row=r, column=1).value
                        _write_person_row(ws_sub, r, sub_vals)

    wb.save(excel_path)
    wb.close()
    return int(stt)


def delete_person(stt, force=False, excel_path=EXCEL_PATH):
    """Xóa nhân đinh theo STT. Cảnh báo nếu nhân đinh đang có con cháu."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Toàn bộ nhân đinh"]

    target_row = None
    target_ma = ""
    children = []
    for r in range(5, ws.max_row + 1):
        s = ws.cell(row=r, column=1).value
        if s is not None and str(s).isdigit() and int(s) == int(stt):
            target_row = r
            target_ma = str(ws.cell(row=r, column=2).value or "").strip()
        if _parse_id_cha(ws.cell(row=r, column=16).value) == int(stt):
            children.append(str(ws.cell(row=r, column=5).value or f"STT {s}"))

    if target_row is None:
        wb.close()
        raise ValueError(f"Không tìm thấy nhân đinh có STT {stt}.")

    if children and not force:
        wb.close()
        raise ValueError(
            f"Nhân đinh này đang là Cha của {len(children)} người ({', '.join(children[:5])}). "
            "Hãy chuyển liên kết con cái trước hoặc xác nhận xóa bắt buộc."
        )

    wb.close()
    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Toàn bộ nhân đinh"]

    if children and force:
        for r in range(5, ws.max_row + 1):
            if _parse_id_cha(ws.cell(row=r, column=16).value) == int(stt):
                ws.cell(row=r, column=16).value = 0

    ws.delete_rows(target_row, 1)

    if target_ma:
        for sname in ("Can 1 - Tổ Cư", "Can 2 - Tổ Khâu", "Can 3 & Can 4"):
            if sname in wb.sheetnames:
                ws_sub = wb[sname]
                for r in range(5, ws_sub.max_row + 1):
                    if str(ws_sub.cell(row=r, column=2).value or "").strip() == target_ma:
                        ws_sub.delete_rows(r, 1)
                        break

    wb.save(excel_path)
    wb.close()
    return True


def add_death(payload, excel_path=EXCEL_PATH):
    """Thêm dòng mới vào Sheet Bảng Kỵ Giỗ."""
    ngay_str = str(payload.get("ngay_str") or "").strip()
    ten = str(payload.get("ten") or "").strip()
    if not ngay_str or not ten:
        raise ValueError("Ngày giỗ Âm lịch và Tên người được giỗ không được trống.")

    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Bảng Kỵ Giỗ"]
    last_row = _find_last_data_row(ws, name_col=3)
    new_stt = (last_row - 4) + 1
    vals = [
        new_stt,
        ngay_str,
        ten,
        str(payload.get("can") or "").strip(),
        str(payload.get("mo") or "").strip(),
        str(payload.get("ghi") or "").strip(),
    ]
    for idx, val in enumerate(vals, start=1):
        cell = ws.cell(row=last_row + 1, column=idx)
        cell.value = val
        cell.font = Font(name="Times New Roman", size=11)
    wb.save(excel_path)
    wb.close()
    return last_row + 1


def update_death(row_idx, payload, excel_path=EXCEL_PATH):
    """Cập nhật dòng kỵ giỗ trong Sheet Bảng Kỵ Giỗ."""
    ngay_str = str(payload.get("ngay_str") or "").strip()
    ten = str(payload.get("ten") or "").strip()
    if not ngay_str or not ten:
        raise ValueError("Ngày giỗ Âm lịch và Tên người được giỗ không được trống.")

    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Bảng Kỵ Giỗ"]
    if row_idx < 5 or row_idx > ws.max_row or not ws.cell(row_idx, 3).value:
        wb.close()
        raise ValueError(f"Không tìm thấy dòng kỵ giỗ ở vị trí {row_idx}.")

    ws.cell(row_idx, 2).value = ngay_str
    ws.cell(row_idx, 3).value = ten
    ws.cell(row_idx, 4).value = str(payload.get("can") or "").strip()
    ws.cell(row_idx, 5).value = str(payload.get("mo") or "").strip()
    ws.cell(row_idx, 6).value = str(payload.get("ghi") or "").strip()
    wb.save(excel_path)
    wb.close()
    return row_idx


def delete_death(row_idx, excel_path=EXCEL_PATH):
    """Xóa dòng kỵ giỗ khỏi Sheet Bảng Kỵ Giỗ."""
    create_backup(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb["Bảng Kỵ Giỗ"]
    if row_idx < 5 or row_idx > ws.max_row or not ws.cell(row_idx, 3).value:
        wb.close()
        raise ValueError(f"Không tìm thấy dòng kỵ giỗ ở vị trí {row_idx}.")
    ws.delete_rows(row_idx, 1)
    wb.save(excel_path)
    wb.close()
    return True


def export_public_html(
    excel_path=EXCEL_PATH,
    template_path=HTML_TEMPLATE_PATH,
    output_path=PUBLIC_HTML_PATH,
    hide_private=False,
):
    """Xuất file HTML tĩnh độc lập sẵn sàng upload lên Internet."""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Không tìm thấy file giao diện gốc {template_path}")

    data = load_all_data(excel_path)
    js_records = []
    for r in data["people"]:
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
            "ngay_mat": r["ngay_mat"],
            "ngay_gio": r["ngay_gio"],
            "cha": r["cha"],
            "me": r["me"],
            "phoi_ngau": r["phoi_ngau"],
            "mo_phan": r["mo_phan"],
            "ghi_chu": r["ghi_chu"],
            "id_cha": r["id_cha"],
            "sdt": "" if hide_private else r["sdt"],
            "facebook": "" if hide_private else r["facebook"],
            "dia_chi": "" if hide_private else r["dia_chi"],
            "nghe_nghiep": r["nghe_nghiep"],
        })

    death_person_map = {}
    for r in data["people"]:
        if r["ngay_gio"]:
            d_r, m_r = parse_lunar_date(r["ngay_gio"])
            if d_r and m_r:
                death_person_map[(d_r, m_r, clean_person_name(r["ho_ten"]))] = r

    js_deaths = []
    for it in data["deaths"]:
        mo_text = f" • Mộ: {it['mo']}" if (it["mo"] and it["mo"] not in it["ghi"]) else ""
        matched = death_person_map.get((it["ngay"], it["thang"], clean_person_name(it["ten"])))
        js_deaths.append({
            "thang": it["thang"],
            "ngay": it["ngay"],
            "ten": it["ten"],
            "can": it["can"],
            "ngay_gio": (matched.get("ngay_gio") if matched else "") or f"{it['ngay']}/{int(it['thang']):02d} ÂL",
            "ngay_mat": matched.get("ngay_mat", "") if matched else "",
            "nam_mat": matched.get("nam_mat", "") if matched else "",
            "ghi": f"{it['ghi']}{mo_text}".strip(" •"),
        })

    records_json = json.dumps(js_records, ensure_ascii=False, indent=2)
    deaths_json = json.dumps(js_deaths, ensure_ascii=False, indent=2)

    with open(template_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    html_content = re.sub(r"const people = \[[\s\S]*?\];", lambda m: f"const people = {records_json};", html_content)
    html_content = re.sub(r"const deathDays = \[[\s\S]*?\];", lambda m: f"const deathDays = {deaths_json};", html_content)
    html_content = re.sub(r"Tổng cộng \d+ Ngày Giỗ", f"Tổng cộng {len(js_deaths)} Ngày Giỗ", html_content)
    html_content = re.sub(r"Lịch Kỵ Giỗ Âm Lịch \(\d+\)", f"Lịch Kỵ Giỗ Âm Lịch ({len(js_deaths)})", html_content)
    html_content = re.sub(r"Tra Cứu Nhân Đinh \(\d+\)", f"Tra Cứu Nhân Đinh ({len(js_records)})", html_content)
    html_content = re.sub(
        r'(<div class="stat-num">)\d+(</div>\s*<div class="stat-label">Tổng số [^<]*</div>)',
        rf"\g<1>{len(js_records)}\g<2>",
        html_content,
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    return {
        "output_path": output_path,
        "people_count": len(js_records),
        "deaths_count": len(js_deaths),
        "hide_private": hide_private,
    }
