"""
Nạp danh sách con cháu đã rà soát (có dấu 'X' ở cột Duyet_Them (X))
từ Danh_Sach_Bo_Sung_Ra_Soat.xlsx vào Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx,
đồng bộ luôn các bản ghi 127..132 từ index.html nếu Excel chưa có,
sau đó xuất lại Gia_Pha_Public.html và index.html.
"""

import json
import pathlib
import re
import shutil
from datetime import datetime
import openpyxl
import genealogy_service as svc


def main():
    # 1. Backup Excel trước khi chỉnh sửa
    bak_dir = pathlib.Path("backups")
    bak_dir.mkdir(exist_ok=True)
    bak = bak_dir / f"Danh_Sach_Nhan_Dinh_Ho_Bach_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    shutil.copy2(svc.EXCEL_PATH, bak)
    print("Backup created:", bak)

    # 2. Đọc dữ liệu 132 người hiện có từ index.html
    idx_path = pathlib.Path("index.html")
    idx_txt = idx_path.read_text(encoding="utf-8")
    idx_people = json.loads(re.search(r"const people = (\[.*?\]);", idx_txt, re.S).group(1))
    idx_death_days_raw = re.search(r"const deathDays = (\[.*?\]);", idx_txt, re.S).group(0)

    # Đồng bộ những người có ở index.html (VD: 127..132) mà Excel local chưa có
    cur_data = svc.load_all_data(svc.EXCEL_PATH)
    xl_stts = {p["stt"] for p in cur_data["people"]}
    synced_base = 0
    for p in sorted(idx_people, key=lambda x: x["stt"]):
        if p["stt"] not in xl_stts:
            payload = {
                "ho_ten": p.get("ten", ""),
                "huy": p.get("huy", ""),
                "gioi_tinh": p.get("gioi_tinh", "Nam"),
                "can": p.get("can", "Can 2"),
                "doi": p.get("doi", "Đời 8"),
                "ma": p.get("ma", ""),
                "id_cha": p.get("id_cha", 0),
                "cha": p.get("cha", ""),
                "me": p.get("me", ""),
                "phoi_ngau": p.get("phoi_ngau", ""),
                "nam_sinh": p.get("nam_sinh", ""),
                "nam_mat": p.get("nam_mat", ""),
                "ngay_mat": p.get("ngay_mat", ""),
                "ngay_gio": p.get("ngay_gio", ""),
                "mo_phan": p.get("mo_phan", ""),
                "sdt": p.get("sdt", ""),
                "facebook": p.get("facebook", ""),
                "dia_chi": p.get("dia_chi", ""),
                "nghe_nghiep": p.get("nghe_nghiep", ""),
                "thu_tu": p.get("thu_tu", ""),
                "ghi_chu": p.get("ghi_chu", ""),
            }
            new_stt = svc.add_person(payload, excel_path=svc.EXCEL_PATH)
            xl_stts.add(new_stt)
            synced_base += 1
    print(f"Synced {synced_base} missing base records from index.html into Excel.")

    # Reload sau khi đồng bộ base
    cur_data = svc.load_all_data(svc.EXCEL_PATH)
    by_stt = {p["stt"]: p for p in cur_data["people"]}

    # 3. Đọc file rà soát
    wb = openpyxl.load_workbook("Danh_Sach_Bo_Sung_Ra_Soat.xlsx")
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    stt_remap = {}  # STT_Goi_Y -> actual new_stt
    added_count = 0
    skipped_count = 0

    for r in rows:
        stt_goi_y = r[0]
        raw_id_cha = int(r[1] or 0)
        can = str(r[3] or "Can 2").strip()
        doi = str(r[4] or "Đời 8").strip()
        ho_ten = str(r[5] or "").strip()
        gioi_tinh = str(r[6] or "Nam").strip()
        phoi_ngau = str(r[7] or "").strip()
        nam_sinh = str(r[8] or "").strip()
        nam_mat = str(r[9] or "").strip()
        ngay_gio = str(r[10] or "").strip()
        noi_o = str(r[11] or "").strip()
        trich_doan = str(r[12] or "").strip()
        duyet = str(r[13] or "").strip().lower()

        if duyet != "x" or not ho_ten:
            skipped_count += 1
            continue

        actual_id_cha = stt_remap.get(raw_id_cha, raw_id_cha)
        parent = by_stt.get(actual_id_cha)
        parent_name = (parent.get("ho_ten") or parent.get("ten") or "") if parent else ""
        cha_name = parent_name
        me_name = ""
        if parent and (parent.get("gioi_tinh") or "").strip().lower() == "nữ":
            me_name = parent_name
            cha_name = (parent.get("phoi_ngau") or "").split(";")[0].split("(")[0].strip()

        payload = {
            "ho_ten": ho_ten,
            "gioi_tinh": gioi_tinh,
            "can": can,
            "doi": doi,
            "id_cha": actual_id_cha,
            "cha": cha_name,
            "me": me_name,
            "phoi_ngau": phoi_ngau,
            "nam_sinh": nam_sinh,
            "nam_mat": nam_mat,
            "ngay_gio": ngay_gio,
            "dia_chi": noi_o,
            "ghi_chu": trich_doan,
        }
        new_stt = svc.add_person(payload, excel_path=svc.EXCEL_PATH)
        if stt_goi_y is not None:
            stt_remap[int(stt_goi_y)] = new_stt
        by_stt[new_stt] = {
            "stt": new_stt,
            "ten": ho_ten,
            "gioi_tinh": gioi_tinh,
            "phoi_ngau": phoi_ngau,
            "can": can,
            "doi": doi,
            "id_cha": actual_id_cha,
        }
        added_count += 1

    print(f"Added {added_count} approved children into Excel (skipped {skipped_count}).")

    # 4. Kết xuất lại Gia_Pha_Public.html & index.html
    svc.export_public_html(excel_path=svc.EXCEL_PATH, output_path=svc.PUBLIC_HTML_PATH, hide_private=False)
    pub_txt = pathlib.Path(svc.PUBLIC_HTML_PATH).read_text(encoding="utf-8")
    # Giữ nguyên deathDays hiện có nếu Excel không có thêm dòng kỵ giỗ mới từ Sheets
    idx_new = re.sub(r"const deathDays = \[.*?\];", lambda _: idx_death_days_raw, pub_txt, count=1, flags=re.S)
    idx_path.write_text(idx_new, encoding="utf-8")

    # 5. Kiểm tra toàn vẹn dữ liệu xuất ra
    final_people = json.loads(re.search(r"const people = (\[.*?\]);", idx_new, re.S).group(1))
    all_stts = {p["stt"] for p in final_people}
    broken_links = [p for p in final_people if (p.get("id_cha") or 0) not in (0, *all_stts)]
    females = sum(1 for p in final_people if (p.get("gioi_tinh") or "").strip().lower() == "nữ")
    males = sum(1 for p in final_people if (p.get("gioi_tinh") or "").strip().lower() == "nam")
    print(f"Final index.html: {len(final_people)} people (Nam: {males}, Nu: {females}), broken id_cha: {len(broken_links)}")
    assert len(broken_links) == 0, f"Found broken id_cha links: {broken_links}"


if __name__ == "__main__":
    main()
