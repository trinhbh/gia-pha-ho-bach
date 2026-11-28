"""
Bóc tách danh sách con ruột (cả Nam và Nữ) được nhắc trong phần
'Tiểu sử, công trạng & Hậu duệ' (ghi_chu) của 132 nhân đinh hiện tại
nhưng chưa có node trên cây phả hệ, xuất ra file Excel rà soát:
    Danh_Sach_Bo_Sung_Ra_Soat.xlsx

An toàn 100%: chỉ ĐỌC index.html và XUẤT file rà soát riêng,
không sửa đổi dữ liệu thật.
"""

import json
import pathlib
import re
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def norm_name(s: str) -> str:
    s = strip_accents(s).lower()
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def compute_next_doi(doi_str: str) -> str:
    d = (doi_str or "").strip()
    if "khởi" in d.lower():
        return "Đời 1"
    m = re.search(r"(\d+)", d)
    if m:
        return f"Đời {int(m.group(1)) + 1}"
    return "Đời 1"


def parse_child_details(segment: str, default_gender: str) -> dict:
    """Trích xuất năm sinh, mất/thọ, ngày giỗ, phối ngẫu, nơi ở từ đoạn văn bản của 1 người con."""
    seg = segment.strip()
    gender = default_gender

    # Năm sinh
    nam_sinh = ""
    m_ns = re.search(
        r"sinh\s+(?:ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+(\d{4})|năm\s+(\d{4}|[A-ZÀ-Ỹ][a-zà-ỹ]+\s+[A-ZÀ-Ỹ][a-zà-ỹ]+)|(\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}))",
        seg,
        re.I,
    )
    if m_ns:
        nam_sinh = next((g for g in m_ns.groups() if g), "").strip()
        # Kèm Can Chi nếu có ngay sau năm
        m_cc = re.search(r"sinh[^.;\n]{0,40}?\b(\d{4})\s*\(([^)]+)\)", seg, re.I)
        if m_cc:
            nam_sinh = f"{m_cc.group(1)} ({m_cc.group(2).strip()})"

    # Năm mất / Thọ / Mất nhỏ
    nam_mat = ""
    if re.search(r"(?:chết|mất)\s+(?:lúc|khi)\s+(?:còn\s+)?nhỏ", seg, re.I):
        nam_mat = "Mất khi còn nhỏ"
    elif re.search(r"mất\s+(?:lúc|khi)\s+chưa\s+có\s+gia\s+thất", seg, re.I):
        nam_mat = "Mất khi chưa có gia thất"
    else:
        m_nm = re.search(r"mất\s+(?:ngày\s+[^,.;\n]+?\s+)?năm\s+(\d{4})(?:\s*\(([^)]+)\))?", seg, re.I)
        m_tho = re.search(r"(?:hưởng\s+)?(?:thọ|hưởng\s+dương)\s+(\d+)\s*tuổi", seg, re.I)
        parts = []
        if m_nm:
            parts.append(f"{m_nm.group(1)}" + (f" ({m_nm.group(2).strip()})" if m_nm.group(2) else ""))
        if m_tho:
            parts.append(f"Thọ {m_tho.group(1)}")
        nam_mat = " - ".join(parts)

    # Ngày giỗ ÂL
    ngay_gio = ""
    m_gio = re.search(
        r"giỗ\s+(?:ngày\s+)?(mồng\s+\d+|\d{1,2})\s+tháng\s+(giêng|chạp|\d{1,2})",
        seg,
        re.I,
    )
    if m_gio:
        d_raw = re.sub(r"mồng\s+", "", m_gio.group(1), flags=re.I).strip()
        m_raw = m_gio.group(2).strip().lower()
        m_map = {"giêng": "01", "chạp": "12"}
        m_num = m_map.get(m_raw, m_raw.zfill(2))
        ngay_gio = f"{d_raw.zfill(2)}/{m_num} ÂL"

    # Phối ngẫu (Chồng / Vợ)
    phoi_ngau = ""
    m_sp = re.search(
        r"(?:lấy\s+chồng(?:\s+là|\s+người[^,.;\n]*)?|chồng\s+là|lấy\s+ông|lấy\s+vợ(?:\s+là)?|vợ\s+là)\s*(?:ông\s+|bà\s+)?([A-ZÀ-Ỹ][^\n.;()]{2,55})",
        seg,
        re.I,
    )
    if m_sp:
        sp = m_sp.group(1).strip().rstrip(",;.")
        # Cắt bỏ phần dư như "có con là...", "làm nông..."
        sp = re.split(r",\s*(?:có\s+\d*\s*con|đã\s+có|làm\s+|hiện|gia\s+đình|nay|sinh\s+được|tốt\s+nghiệp)", sp, flags=re.I)[0].strip()
        if sp and not re.match(r"^(người|cùng|ở|xã)\b", sp, re.I):
            phoi_ngau = sp

    # Nơi ở hiện tại
    noi_o = ""
    m_addr = re.search(
        r"(?:sống\s+ở|định\s+cư\s+ở|công\s+tác\s+(?:và\s+sống\s+)?ở|nhà\s+(?:hiện\s+)?ở|đang\s+ở)\s+([^\n.;]{3,65})",
        seg,
        re.I,
    )
    if m_addr:
        noi_o = m_addr.group(1).strip().rstrip(",;.")

    return {
        "gioi_tinh": gender,
        "nam_sinh": nam_sinh,
        "nam_mat": nam_mat,
        "ngay_gio": ngay_gio,
        "phoi_ngau": phoi_ngau,
        "noi_o": noi_o,
    }


def get_descendant_norms(parent_stt: int, all_people: list) -> set:
    """Lấy toàn bộ tên chuẩn hóa (và tên húy) của con, cháu, chắt đã có dưới parent_stt."""
    by_parent = {}
    for p in all_people:
        by_parent.setdefault(p.get("id_cha") or 0, []).append(p)
    result = set()
    stack = list(by_parent.get(parent_stt, []))
    while stack:
        node = stack.pop()
        result.add(norm_name(node.get("ten", "")))
        if node.get("huy"):
            result.add(norm_name(node["huy"]))
        # Thêm biến thể bỏ ngoặc / giữ ngoặc
        m_par = re.search(r"\(([^)]+)\)", node.get("ten", ""))
        if m_par:
            result.add(norm_name(m_par.group(1)))
            base = re.sub(r"\([^)]*\)", "", node.get("ten", "")).strip()
            if base:
                result.add(norm_name(base))
        stack.extend(by_parent.get(node["stt"], []))
    return result


def extract_children_from_note(parent: dict, all_people: list) -> list:
    gc = (parent.get("ghi_chu") or "").strip()
    if not gc:
        return []

    # Cắt bỏ phần "Khấn:" nếu có
    gc_main = re.split(r"\n\s*Khấn\s*:", gc, flags=re.I)[0]

    parent_raw = parent.get("ten", "")
    parent_norms = {norm_name(parent_raw)}
    if parent.get("huy"):
        parent_norms.add(norm_name(parent["huy"]))
    m_ppar = re.search(r"\(([^)]+)\)", parent_raw)
    if m_ppar:
        parent_norms.add(norm_name(m_ppar.group(1)))
        # Nếu trong ngoặc chỉ là tên ngắn như "(Tam)" -> ghép họ đệm của cha
        short_alias = m_ppar.group(1).strip()
        if " " not in short_alias and len(parent_raw.split()) >= 2:
            parent_norms.add(norm_name(" ".join(parent_raw.split()[:2]) + " " + short_alias))

    grand_norm = norm_name(parent.get("cha", ""))
    spouse_norms = {norm_name(s) for s in re.split(r"[;,]", parent.get("phoi_ngau") or "") if s.strip()}

    lines = [ln.strip() for ln in gc_main.splitlines() if ln.strip()]

    # Tìm index dòng bắt đầu liệt kê con cái (nếu có)
    child_start_idx = 0
    for idx, ln in enumerate(lines):
        if re.search(
            r"(sinh\s+được|sinh\s+(?:một|hai|ba|bốn|năm|\d+)\s+(?:người\s+)?con|con\s+(?:của\s+)?ông\s*bà|con\s+ông|gồm\s+có|có\s+\d+\s+con|bà\s+(?:nhất|hai|ba|cả).*gồm\s+có)",
            ln,
            re.I,
        ):
            child_start_idx = idx
            break

    candidate_text = "\n".join(lines[child_start_idx:])

    name_pat = re.compile(
        r"(?:(Ông|Bà|Con\s+gái|Con\s+trai|Trưởng\s+nữ|Thứ\s+nữ)\s+)?"
        r"\b(Bạch\s+(?:Thị\s+[A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){0,2}|[A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){1,3}))\b"
    )

    matches = list(name_pat.finditer(candidate_text))
    if not matches:
        return []

    descendant_norms = get_descendant_norms(parent["stt"], all_people)

    results = []
    seen_norms = set()

    for i, m in enumerate(matches):
        prefix = (m.group(1) or "").strip()
        raw_name = re.sub(r"\s+", " ", m.group(2)).strip()
        # Chỉ cắt đuôi nếu là từ viết thường (sinh, mất, lấy...) hoặc từ khóa chức danh/thời điểm
        raw_name = re.sub(r"\s+(?:sinh|mất|lấy|chết|tốt|hiện|nay|làm|giỗ|thọ|đi|ở|là|và|có|đã|từ|khi|sau|trước|vợ|chồng|con|đang|tên|húy)\b.*$", "", raw_name)
        n_norm = norm_name(raw_name)

        if not n_norm or len(n_norm.split()) < 2:
            continue
        if n_norm in parent_norms or n_norm == grand_norm or n_norm in spouse_norms:
            continue

        start = m.start()
        line_start = candidate_text.rfind("\n", 0, start) + 1
        ctx_same_line = candidate_text[line_start:start].lower()
        ctx_before = candidate_text[max(0, start - 90):start].lower()

        if re.search(r"(tổ\s+ông|con\s+trai\s+(?:cả|thứ|út)\s+ông|con\s+ông|lời\s+dặn\s+của\s+tổ|nhắc\s+lại|phụng\s+tự|cháu\s+là)\s*$", ctx_before):
            continue

        # Bỏ qua nếu nằm trong cụm liệt kê cháu nội của một người con (khi đã có dòng dẫn riêng của người con)
        if line_start > 0 and re.search(r"(?:sinh\s+(?:được\s+)?(?:\d+|một|hai|ba|bốn|năm)\s+(?:người\s+)?con(?:\s+gái|\s+trai)?(?:\s+là)?|có\s+(?:\d+|một|hai|ba|bốn|năm)\s+con(?:\s+gái|\s+trai)?(?:\s+là)?|có\s+con\s+là)\s*:?\s*(?:[^.\n]*[;,]\s*)?$", ctx_same_line):
            continue

        end = matches[i + 1].start() if i + 1 < len(matches) else len(candidate_text)
        seg_after = candidate_text[m.start():min(end, m.start() + 350)]
        if re.match(r"^[^.;\n]{0,50}phụng\s+tự", candidate_text[m.end():m.end() + 60], re.I):
            continue

        if n_norm in descendant_norms:
            continue
        if n_norm in seen_norms:
            continue
        seen_norms.add(n_norm)

        is_female = (
            "thị" in raw_name.lower().split()
            or any(w in prefix.lower() for w in ("bà", "nữ", "con gái"))
            or bool(re.search(r"\b(?:lấy\s+chồng|chồng\s+là|cô\s+" + re.escape(raw_name.split()[-1].lower()) + r")\b", seg_after, re.I))
        )
        if re.search(r"^(?:Bạch\s+Thùy|Bạch\s+Thủy|Bạch\s+Trà|Bạch\s+Hồng|Bạch\s+Khánh|Bạch\s+Mai|Bạch\s+Phương\s+Thảo|Bạch\s+Thảo|Bạch\s+Tuệ)", raw_name, re.I):
            is_female = True
        gender = "Nữ" if is_female else "Nam"

        details = parse_child_details(seg_after, gender)
        results.append({
            "parent": parent,
            "ho_ten": raw_name,
            **details,
            "trich_doan": seg_after.strip().replace("\n", " ")[:280],
        })

    return results


def extract_grandchildren_pass2(people: list, pass1_items: list) -> list:
    """
    Lượt 2: Bóc tách thế hệ cháu nội (đời 8, đời 9...) được ghi gộp trong tiểu sử của ông nội
    (ví dụ: 'Bạch Hưng Thư và Tưởng Thị Linh có 2 con là: Bạch Hồng Nhung...; Bạch Hưng Dũng...').
    Gắn vào node cha tương ứng (đã có trong people hoặc vừa tạo ở Lượt 1).
    """
    next_stt_base = max((p.get("stt") or 0) for p in people) + 1
    # Tạo map các node con (stt, ten, can, doi) theo parent_stt
    virtual_nodes = []
    for p in people:
        virtual_nodes.append({
            "stt": p["stt"],
            "id_cha": p.get("id_cha") or 0,
            "ten": p["ten"],
            "can": p["can"],
            "doi": p["doi"],
        })
    for idx, item in enumerate(pass1_items):
        p = item["parent"]
        virtual_nodes.append({
            "stt": next_stt_base + idx,
            "id_cha": p["stt"],
            "ten": item["ho_ten"],
            "can": p["can"],
            "doi": compute_next_doi(p["doi"]),
        })

    by_parent = {}
    for vn in virtual_nodes:
        by_parent.setdefault(vn["id_cha"], []).append(vn)

    name_pat = re.compile(
        r"\b(Bạch\s+(?:Thị\s+[A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){0,2}|[A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){1,3}))\b"
    )
    existing_all_norms = {norm_name(vn["ten"]) for vn in virtual_nodes}

    pass2_items = []
    for p in people:
        gc = (p.get("ghi_chu") or "").strip()
        if not gc:
            continue
        kids_of_p = by_parent.get(p["stt"], [])
        if not kids_of_p:
            continue

        lines = [ln.strip() for ln in gc.splitlines() if ln.strip()]
        for ln_idx, ln in enumerate(lines):
            if ln_idx == 0:
                continue  # Dòng đầu là giới thiệu chính p
            # Tìm cụm liệt kê con của một người con trong dòng
            m_sub = re.search(
                r"(?:sinh\s+(?:được\s+)?(?:\d+|một|hai|ba|bốn|năm)\s+(?:người\s+)?con(?:\s+gái|\s+trai)?(?:\s+là)?|có\s+(?:\d+|một|hai|ba|bốn|năm)\s+con(?:\s+gái|\s+trai)?(?:\s+là)?|có\s+con\s+là)\s*:?\s*(.+)$",
                ln,
                re.I,
            )
            if not m_sub:
                continue

            prefix_part = ln[:m_sub.start()]
            list_part = m_sub.group(1)
            # Xác định sub_parent trong kids_of_p dựa trên tên đầy đủ hoặc tên gọi cuối xuất hiện trong prefix_part
            sub_parent = None
            best_pos = -1
            for k in kids_of_p:
                k_full = k["ten"]
                k_last = k_full.split()[-1]
                pos = prefix_part.lower().rfind(k_full.lower())
                if pos == -1 and len(k_last) >= 2:
                    m_last = re.search(r"\b" + re.escape(k_last) + r"\b", prefix_part, re.I)
                    if m_last:
                        pos = m_last.start()
                if pos > best_pos:
                    best_pos = pos
                    sub_parent = k

            if not sub_parent:
                continue

            sub_existing = {norm_name(c["ten"]) for c in by_parent.get(sub_parent["stt"], [])}
            sub_matches = list(name_pat.finditer(list_part))
            for i, sm in enumerate(sub_matches):
                raw_name = re.sub(r"\s+", " ", sm.group(1)).strip()
                raw_name = re.sub(r"\s+(?:sinh|mất|lấy|chết|tốt|hiện|nay|làm|giỗ|thọ|đi|ở|là|và|có|đã|từ|khi|sau|trước|vợ|chồng|con|đang|tên|húy)\b.*$", "", raw_name)
                n_norm = norm_name(raw_name)
                if not n_norm or len(n_norm.split()) < 2:
                    continue
                if n_norm == norm_name(sub_parent["ten"]) or n_norm in sub_existing or n_norm in existing_all_norms:
                    continue

                end = sub_matches[i + 1].start() if i + 1 < len(sub_matches) else len(list_part)
                seg_after = list_part[sm.start():min(end, sm.start() + 260)]
                is_female = (
                    "thị" in raw_name.lower().split()
                    or bool(re.search(r"con\s+gái\s+là\s+" + re.escape(raw_name.split()[0].lower()), ln, re.I))
                    or bool(re.search(r"^(?:Bạch\s+Thùy|Bạch\s+Thủy|Bạch\s+Trà|Bạch\s+Hồng|Bạch\s+Khánh|Bạch\s+Mai|Bạch\s+Phương\s+Thảo|Bạch\s+Thảo|Bạch\s+Tuệ)", raw_name, re.I))
                )
                gender = "Nữ" if is_female else "Nam"
                details = parse_child_details(seg_after, gender)
                existing_all_norms.add(n_norm)
                sub_existing.add(n_norm)
                pass2_items.append({
                    "parent": sub_parent,
                    "ho_ten": raw_name,
                    **details,
                    "trich_doan": f"[Từ ghi chú {p['ten']}] " + seg_after.strip()[:240],
                })

    return pass2_items


def main():
    idx_txt = pathlib.Path("index.html").read_text(encoding="utf-8")
    people = json.loads(re.search(r"const people = (\[.*?\]);", idx_txt, re.S).group(1))

    all_extracted = []
    for p in people:
        kids = extract_children_from_note(p, people)
        all_extracted.extend(kids)

    pass2_items = extract_grandchildren_pass2(people, all_extracted)
    all_extracted.extend(pass2_items)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Danh Sach Bo Sung Ra Soat"

    headers = [
        "STT_Goi_Y",
        "ID_CHA",
        "Ten_Cha_Tren_Cay",
        "Can_Chi",
        "Doi_Con",
        "Ho_va_Ten_Con",
        "Gioi_Tinh",
        "Phoi_Ngau",
        "Nam_Sinh",
        "Nam_Mat_Tho",
        "Ngay_Gio_AL",
        "Noi_O_Hien_Tai",
        "Trich_Doan_Van_Ban_Goc",
        "Duyet_Them (X)",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="7C2D12", end_color="7C2D12", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    next_stt = max((p.get("stt") or 0) for p in people) + 1
    female_cnt = 0
    male_cnt = 0

    for idx, item in enumerate(all_extracted):
        p = item["parent"]
        p_name = p.get("ten") or p.get("ho_ten") or ""
        if item["gioi_tinh"] == "Nữ":
            female_cnt += 1
        else:
            male_cnt += 1
        ws.append([
            next_stt + idx,
            p["stt"],
            f"{p_name} ({p['can']} - {p['doi']})",
            p["can"],
            compute_next_doi(p["doi"]),
            item["ho_ten"],
            item["gioi_tinh"],
            item["phoi_ngau"],
            item["nam_sinh"],
            item["nam_mat"],
            item["ngay_gio"],
            item["noi_o"],
            item["trich_doan"],
            "X",
        ])

    widths = [10, 10, 28, 12, 10, 22, 10, 28, 18, 22, 14, 26, 55, 14]
    for col_idx, w in enumerate(widths, 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = w

    out_path = pathlib.Path("Danh_Sach_Bo_Sung_Ra_Soat.xlsx")
    wb.save(out_path)

    print(f"Total extracted missing children: {len(all_extracted)} (Nu: {female_cnt}, Nam: {male_cnt})")
    print(f"Saved review Excel: {out_path.resolve()}")

    # Print sample for Bach Hung Yem (#54 or wherever he is)
    yem_items = [x for x in all_extracted if "yêm" in x["parent"]["ten"].lower()]
    print(f"\nSample for Bach Hung Yem ({len(yem_items)}):")
    for x in yem_items:
        print(f"  - {x['ho_ten']} ({x['gioi_tinh']}) | Phoi ngau: {x['phoi_ngau']} | Mat: {x['nam_mat']} | Gio: {x['ngay_gio']}")


if __name__ == "__main__":
    main()
