# -*- coding: utf-8 -*-
"""
Script tái tạo đồng bộ cả Excel và Word dựa 100% vào văn bản thực tế trong sách.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUTPUT_DIR = r"d:\python\gia pha ho Bach"
EXCEL_PATH = os.path.join(OUTPUT_DIR, "Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx")
DOCX_PATH = os.path.join(OUTPUT_DIR, "Gia_Pha_Ho_Bach.docx")

# Load records from update_excel.py
import sys
sys.path.append(OUTPUT_DIR)
from update_excel import records

# -------------------------------------------------------------
# 1. TẠO FILE EXCEL HOÀN CHỈNH
# -------------------------------------------------------------
wb = openpyxl.Workbook()

font_title = Font(name="Times New Roman", size=15, bold=True, color="1F497D")
font_subtitle = Font(name="Times New Roman", size=11, italic=True, color="595959")
font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Times New Roman", size=11, color="000000")
font_bold = Font(name="Times New Roman", size=11, bold=True, color="000000")

fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF')
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

cols_def = [
    ("STT", 6, align_center),
    ("Mã ĐD", 12, align_center),
    ("Can", 14, align_center),
    ("Đời", 10, align_center),
    ("Họ và Tên", 22, align_left),
    ("Tên Húy / Biệt danh", 18, align_left),
    ("Giới tính", 10, align_center),
    ("Năm sinh", 18, align_center),
    ("Năm mất / Thọ", 22, align_center),
    ("Ngày giỗ ÂL", 14, align_center),
    ("Cha", 24, align_left),
    ("Mẹ", 20, align_left),
    ("Phối ngẫu", 32, align_left),
    ("Mộ phần", 22, align_left),
    ("Công trạng / Ghi chú (Theo văn bản)", 45, align_left),
]

def setup_ws(ws, title):
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols_def))
    c1 = ws.cell(row=1, column=1, value=title)
    c1.font = font_title
    c1.alignment = align_center

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols_def))
    c2 = ws.cell(row=2, column=1, value="TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2 - BẠCH TỘC ĐẠI TÔN (HƯNG PHÚ, HƯNG NGUYÊN)")
    c2.font = font_subtitle
    c2.alignment = align_center

    for c_idx, (c_name, c_width, c_align) in enumerate(cols_def, start=1):
        cell = ws.cell(row=4, column=c_idx, value=c_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = c_width

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[4].height = 26

def write_rows(ws, data):
    for r_idx, r in enumerate(data, start=5):
        vals = [
            r_idx - 4, r["ma"], r["can"], r["doi"], r["ho_ten"], r["huy"],
            r["gioi_tinh"], r["nam_sinh"], r["nam_mat"], r["ngay_gio"],
            r["cha"], r["me"], r["phoi_ngau"], r["mo_phan"], r["ghi_chu"]
        ]
        is_zebra = (r_idx % 2 == 0)
        for c_idx, val in enumerate(vals, start=1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font = font_bold if c_idx == 5 else font_body
            c.border = thin_border
            c.alignment = cols_def[c_idx-1][2]
            if is_zebra:
                c.fill = fill_zebra
        ws.row_dimensions[r_idx].height = 22

# Sheet 1: Toàn bộ nhân đinh
ws_all = wb.active
ws_all.title = "Toàn bộ nhân đinh"
setup_ws(ws_all, "DANH BẠ NHÂN ĐINH DÒNG HỌ BẠCH (CHUẨN THEO VĂN BẢN GỐC)")
write_rows(ws_all, records)

# Sheet 2: Can 1
ws_c1 = wb.create_sheet(title="Can 1 - Tổ Cư")
setup_ws(ws_c1, "DANH BẠ CAN 1 - TỔ BẠCH HƯNG CƯ (ĐỜI 3 ĐẾN ĐỜI 8)")
c1_data = [r for r in records if r["can"] in ["Can 1", "Khởi phát", "Tiểu Chi 1"]]
write_rows(ws_c1, c1_data)

# Sheet 3: Can 2
ws_c2 = wb.create_sheet(title="Can 2 - Tổ Khâu")
setup_ws(ws_c2, "DANH BẠ CAN 2 - TỔ BẠCH HƯNG KHÂU (ĐỜI 3 ĐẾN ĐỜI 8)")
c2_data = [r for r in records if r["can"] in ["Can 2", "Khởi phát", "Tiểu Chi 1"]]
write_rows(ws_c2, c2_data)

# Sheet 4: Can 3 & Can 4
ws_c34 = wb.create_sheet(title="Can 3 & Can 4")
setup_ws(ws_c34, "DANH BẠ CAN 3 & CAN 4 - TỔ BẠCH HƯNG NGHI & TỔ BẠCH HƯNG KẾ")
c34_data = [r for r in records if r["can"] in ["Can 3", "Can 4", "Khởi phát", "Tiểu Chi 1"]]
write_rows(ws_c34, c34_data)

# Sheet 5: Kỵ giỗ
gio_list = [
    (1, "10 tháng Giêng", "Cụ bà Phạm Thị Biềng", "Đời 2", "Xứ Thượng Phù", "Vợ Cụ Bạch Hưng Lưu"),
    (2, "23 tháng Giêng", "Ông Bạch Hưng Hợi", "Can 2 - Đời 5", "Nghĩa Quang, Nghĩa Đàn", "Mất lúc 51 tuổi. Con Cụ Xu"),
    (3, "26 tháng Giêng", "Cụ bà Nguyễn Thị Tư", "Can 2 - Đời 5", "Nghĩa Quang, Nghĩa Đàn", "Vợ ông Bạch Hưng Hợi. Thọ 83 tuổi"),
    (4, "28 tháng Giêng", "Cụ bà Bùi Thị Hàng Hai", "Khởi phát", "Xứ Khe Dài", "Vợ hai Cụ Bạch Hưng Hóa"),
    (5, "30 tháng Giêng", "Bà Hồ Thị Hảo", "Can 2 - Đời 6", "Núi Khe Xứ", "Vợ đầu Liệt sỹ Bạch Hưng Đình (1922-1951)"),
    (6, "30 tháng Giêng", "Ông Bạch Hưng Tứ", "Can 2 - Đời 6", "Nghĩa Thịnh, Nghĩa Đàn", "Con thứ 2 ông Bạch Hưng Hợi (1923-1999)"),
    (7, "04 tháng Hai", "Cụ bà Phạm Thị Đích", "Can 2 - Đời 6", "Núi Khe xứ", "Vợ ông Bạch Hưng Lực. Thọ 83 tuổi (1932-2014)"),
    (8, "08 tháng Hai", "Cụ bà Nguyễn Thị Máy", "Can 1 - Đời 3", "Cồn Giữa xứ", "Vợ cả Cụ Bạch Hưng Cư"),
    (9, "10 tháng Hai", "Ông Bạch Hưng Mận", "Can 1 - Đời 6", "Nghĩa Đức, Nghĩa Đàn", "Thọ 68 tuổi (1935-2002). Con ông Chế (Tam)"),
    (10, "15 tháng Hai", "Cụ bà Ngô Thị Đôi", "Can 3 - Đời 5", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Nguyễn. Thọ 66 tuổi"),
    (11, "15 tháng Hai", "Cụ Bạch Hưng Hiến (Đầu phiên)", "Can 3 - Đời 6", "Núi Khe xứ", "Thọ 82 tuổi (1906-1986). Con cả Cụ Khuyến"),
    (12, "22 tháng Hai", "Cụ Bạch Hưng Tam (Chắt Mãi)", "Can 2 - Đời 5", "Núi Khe Xứ", "Thọ 60 tuổi (1891-1950). Con cả Cụ Xu"),
    (13, "29 tháng Hai", "Cụ Bạch Hưng Lưu (húy Túc)", "Đời 2", "Xứ Thượng Phù", "Thọ 76 tuổi (1798-1873). Khởi tổ 4 Can"),
    (14, "01 tháng Ba", "Cụ Bạch Hưng Thiết (húy Khải)", "Can 1 - Đời 4", "Khe Xứ", "Con Cụ Bạch Hưng Cư"),
    (15, "03 tháng Ba", "Ông Bạch Hưng Vỹ", "Can 2 - Đời 6", "Núi Khe xứ", "Con thứ 2 ông Yêm, mất không con"),
    (16, "04 tháng Ba", "Ông Bạch Hưng Minh", "Can 3 - Đời 7", "", "Thọ 47 tuổi (1948-1993). Kế toán trưởng Cty Vật tư NN"),
    (17, "07 tháng Ba", "Liệt sỹ Bạch Hưng Xuân", "Can 2 - Đời 6", "Nghĩa trang QG Tuyên Hóa, Quảng Bình", "Hy sinh 3/1966 tuổi 34. Con Cụ Em (Con)"),
    (18, "12 tháng Ba", "Bạch Hưng Hưng (húy Kiến)", "Can 3 - Đời 6", "Núi Khe xứ", "Mất lúc 14 tuổi. Con Cụ Khuyến"),
    (19, "15 tháng Ba", "Ông Bạch Hưng Em (Bâu)", "Can 1 - Đời 5", "Xứ Khe", "Con Cụ Tơn"),
    (20, "15 tháng Ba", "Cụ bà Hồ Thị Nhiên", "Can 2 - Đời 5", "Thồ Lô xứ", "Vợ Cụ Bạch Sĩ Xum (Cố Hòe)"),
    (21, "17 tháng Ba", "Ông Bạch Hưng Dĩ", "Can 2 - Đời 5", "Núi Khe xứ (Rú sau)", "Thọ 61 tuổi. Con thứ 2 Cụ Câu"),
    (22, "19 tháng Ba", "Cụ bà Hồ Thị Út", "Can 2 - Đời 4", "Thồ Lô xứ", "Vợ Cụ Bạch Hưng Thưu"),
    (23, "23 tháng Ba", "Ông Bạch Hưng Hai", "Can 1 - Đời 5", "Núi Khe xứ", "Con Cụ Ngân"),
    (24, "23 tháng Ba", "Ông Bạch Hưng Nhạ", "Can 1 - Đời 6", "Nghĩa Đức, Nghĩa Đàn", "Thọ 76 tuổi (1923-1998). Con thứ 2 ông Chế"),
    (25, "27 tháng Ba", "Ông Bạch Hưng Bốn (Uyển)", "Can 2 - Đời 6", "Núi Khe xứ", "Thọ 64 tuổi (1930-1993). Con thứ 3 ông Dĩ"),
    (26, "28 tháng Ba", "Cụ bà Nguyễn Quý Nương (Vân)", "Can 1 - Đời 3", "Xứ Đồng Trọt Bàn", "Vợ Cụ Cư. Thọ 85 tuổi"),
    (27, "28 tháng Ba", "Cụ Bạch Hưng Em (húy Con)", "Can 2 - Đời 5", "Núi Khe xứ", "Thọ 82 tuổi (1900-1981). Con thứ 3 Cụ Câu"),
    (28, "01 tháng Tư", "Bà Võ Thị Chiên", "Can 2 - Đời 7", "Đức Trọng, Lâm Đồng", "Vợ ông Bạch Hưng Huy. Mất 2012 tuổi 50"),
    (29, "05 tháng Tư", "Cụ Bạch Hưng Thấu", "Can 2 - Đời 4", "Kẹt Đá xứ", "Mất khi chưa có gia thất. Con cả Cụ Khâu"),
    (30, "10 tháng Tư", "Cụ Bạch Hưng Nghi (Thủ chỉ)", "Can 3 - Đời 3", "Núi Khe xứ", "Khởi tổ Can 3. Thọ 89 tuổi"),
    (31, "14 tháng Tư", "Cụ BẠCH HƯNG HÓA", "Khởi phát", "Xứ Cồn Xanh", "Tổ Khởi phát Trung Chi 2. Thọ 73 tuổi"),
    (32, "15 tháng Tư", "Cụ bà Phạm Thị Đị", "Can 3 - Đời 6", "Núi Khe xứ", "Vợ Cụ Kiểm Nhưng"),
    (33, "07 tháng Năm", "Cụ bà Phạm Thị Phòng", "Can 1 - Đời 4", "", "Vợ Cụ Bạch Hưng Tơn"),
    (34, "09 tháng Năm", "Cụ bà Nguyễn Thị Lù", "Can 2 - Đời 5", "Núi Khe xứ", "Vợ ông Bạch Hưng Dĩ"),
    (35, "11 tháng Năm", "Ông Bạch Hưng Niêm", "Can 2 - Đời 6", "Hố Đót, Rú Trước", "Con thứ 3 ông Hợi (1930-2011)"),
    (36, "15 tháng Năm", "Ông Bạch Hưng Niêm", "Can 3 - Đời 6", "Núi Khe xứ", "Mất khi chưa lấy vợ. Con Cụ Nguyễn"),
    (37, "19 tháng Năm", "Ông Bạch Hưng Phòng", "Can 2 - Đời 5", "Núi Khe Xứ", "Mất chưa có gia thất. Con cả Cụ Câu"),
    (38, "24 tháng Năm", "Cụ Bạch Hưng Cường", "Can 2 - Đời 6", "Đức Trọng, Lâm Đồng", "Mất 24/05/2016 (Bính Thân) tại Lâm Đồng. Thọ 97 tuổi"),
    (39, "28 tháng Năm", "Cụ Bạch Hưng Khâm", "Đời 2", "Xứ Cồn Vải", "Mất năm 20 tuổi. Con trưởng Cụ Tạo"),
    (40, "01 tháng Sáu", "Cụ Bạch Hưng Xu (húy Mâu)", "Can 2 - Đời 4", "Thồ Lô xứ", "Con thứ 3 Cụ Khâu (1854)"),
    (41, "12 tháng Sáu", "Cụ Bạch Hưng Tơn", "Can 1 - Đời 4", "", "Thọ 60 tuổi. Con Cụ Cư"),
    (42, "13 tháng Sáu", "Ông Bạch Hưng Nậy", "Can 2 - Đời 6", "Nghĩa Đức, Nghĩa Đàn", "Con cả ông Hợi (1918-1993)"),
    (43, "25 tháng Sáu", "Cụ bà Nguyễn Thị Lam", "Can 2 - Đời 6", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Cường. Hưởng dương 39 tuổi (1968)"),
    (44, "25 tháng Sáu", "Ông Bạch Hưng Lực", "Can 2 - Đời 6", "Núi Khe xứ", "Thọ 79 tuổi (1931-2009). Con cả Cụ Em (Con)"),
    (45, "28 tháng Sáu", "Cụ Bạch Hưng Phù (Tư Lang)", "Can 2 - Đời 4", "Khe Xứ", "Thọ 51 tuổi (1862-1912). Không lập gia thất"),
    (46, "28 tháng Sáu", "Ông Bạch Hưng Sáu", "Can 3 - Đời 6", "Núi Khe xứ", "Mất ở tuổi 35. Con Cụ Nguyễn"),
    (47, "30 tháng Sáu", "Cụ Bạch Hưng Ngân", "Can 1 - Đời 4", "Rú Sau (cây đa Bò Mẹp)", "Con Cụ Cư"),
    (48, "30 tháng Sáu", "Ông Bạch Hưng Cháu", "Can 1 - Đời 6", "Núi Khe Xứ", "Con cả ông Chế (1923-2014)"),
    (49, "30 tháng Sáu", "Cụ bà Nguyễn Thị Ngụ (Thông)", "Can 3 - Đời 3", "Núi Khe xứ", "Mất năm 36 tuổi. Vợ hai Cụ Nghi"),
    (50, "03 tháng Bảy", "Liệt sỹ Bạch Hưng Đình", "Can 2 - Đời 6", "Núi Khe Xứ", "Chủ tịch xã Hưng Phú, hy sinh 27/8/1968 (Mậu Thân)"),
    (51, "11 tháng Bảy", "Cụ Bạch Hưng Truyện (Lý trưởng)", "Can 3 - Đời 4", "Núi Khe xứ", "Mất năm 28 tuổi (bị hại). Con Cụ Nghi"),
    (52, "12 tháng Bảy", "Ông Bạch Hưng Thành", "Can 3 - Đời 7", "Núi Khe xứ", "Thọ 49 tuổi. Huân chương Kháng chiến Hạng Nhất"),
    (53, "15 tháng Bảy", "Cụ bà Hồ Thị Niêm", "Can 1 - Đời 4", "Núi Khe xứ", "Vợ cả Cụ Bạch Hưng Ngân"),
    (54, "26 tháng Bảy", "Cụ bà Hồ Thị Em", "Can 3 - Đời 5", "Núi Khe xứ", "Thọ 73 tuổi. Vợ Cụ Bạch Hưng Khuyến"),
    (55, "01 tháng Tám", "Cụ bà Cao Thị Hoét", "Can 3 - Đời 6", "Núi Khe xứ", "Mất năm 40 tuổi. Vợ Cụ Đầu phiên Hiến"),
    (56, "02 tháng Tám", "Cụ Bạch Hưng Khâu (húy Hoài)", "Can 2 - Đời 3", "Núi Khe xứ", "Khởi tổ Can 2. Thọ 77 tuổi (1821-1897)"),
    (57, "06 tháng Tám", "Ông Bạch Hưng Em", "Can 2 - Đời 5", "Thồ Lô xứ", "Con thứ 4 Cụ Xu. Mất chưa có gia thất"),
    (58, "07 tháng Tám", "Cụ bà Bùi Thị Hàng Nhất", "Khởi phát", "Xứ Cồn Xanh", "Vợ cả Cụ Bạch Hưng Hóa"),
    (59, "07 tháng Tám", "Cụ bà Bùi Thị Muồng", "Can 2 - Đời 3", "Xứ Rộc Mâm", "Vợ cả Cụ Bạch Hưng Khâu"),
    (60, "11 tháng Tám", "Cụ bà Nguyễn Thị Hàng Hai", "Khởi phát", "Xứ Khe Dài", "Vợ ba Cụ Bạch Hưng Hóa"),
    (61, "11 tháng Tám", "Cụ Bạch Sĩ Xum (Cố Hòe)", "Can 2 - Đời 5", "Thồ Lô xứ", "Con Cụ Bạch Hưng Thưu"),
    (62, "11 tháng Tám", "Ông Bạch Hưng Nậy", "Can 2 - Đời 6", "Núi Khe xứ", "Con cả ông Dĩ, mất năm ~30 tuổi"),
    (63, "20 tháng Tám", "Cụ Bạch Hưng Hai (Kiểm Nhưng)", "Can 3 - Đời 6", "Núi Khe xứ", "Thọ 54 tuổi. Con cả Cụ Nguyễn"),
    (64, "28 tháng Tám", "Cụ bà Phạm Thị Hoạt", "Can 2 - Đời 3", "Xứ Cồn Xanh", "Vợ hai Cụ Khâu. Thọ 62 tuổi"),
    (65, "06 tháng Chín", "Cụ Bạch Hưng Khuyến (Cai xã)", "Can 3 - Đời 5", "Núi Khe xứ", "Thọ 66 tuổi. Con thứ Cụ Truyện"),
    (66, "17 tháng Chín", "Cụ bà Nguyễn Thị Em (húy Tảo)", "Can 1 - Đời 4", "Rú sau", "Vợ hai Cụ Bạch Hưng Ngân"),
    (67, "23 tháng Chín", "Bà Phạm Thị Nhỏ", "Can 2 - Đời 6", "Nghĩa Thịnh", "Vợ ông Bạch Hưng Tứ"),
    (68, "24 tháng Chín", "Cụ bà Ngô Thị Tư", "Can 2 - Đời 4", "Núi Khe Xứ", "Vợ Cụ Bạch Hưng Câu"),
    (69, "28 tháng Chín", "Cụ BẠCH TRẦN TẠO", "Đời 1", "Xứ Tuần Tượng", "Khởi tổ Tiểu Chi 1. Thọ 30 tuổi (1772-1801)"),
    (70, "28 tháng Chín", "Cụ Bạch Hưng Câu", "Can 2 - Đời 4", "Núi Khe Xứ", "Thọ 31 tuổi (1864-1894). Con thứ 5 Cụ Khâu"),
    (71, "01 tháng Mười", "Ông Bạch Hưng Ý", "Can 1 - Đời 4", "Núi Khe xứ", "Con thứ 3 Cụ Cư, chết đuối sông Hiếu"),
    (72, "08 tháng Mười", "Cụ bà Phạm Thị Hàng Nhất", "Đời 1", "Xứ Cồn Xanh", "Thọ 76 tuổi. Vợ Cụ Bạch Trần Tạo"),
    (73, "09 tháng Mười", "Cụ Bạch Hưng Cư (húy Khuơ)", "Can 1 - Đời 3", "Núi Khe xứ", "Khởi tổ Can 1. Sinh 1819"),
    (74, "11 tháng Mười", "Cụ bà Nguyễn Thị Tứ", "Can 1 - Đời 5", "Nghĩa Đức", "Thọ 88 tuổi (1900-1987). Vợ Cụ Bạch Hưng Chế"),
    (75, "25 tháng Mười", "Cụ Bạch Hưng Thưu", "Can 2 - Đời 4", "Thồ Lô xứ", "Con thứ 2 Cụ Khâu"),
    (76, "14 tháng Mười Một", "Cụ bà Hồ Thị Phong", "Can 3 - Đời 3", "Núi Khe xứ", "Thọ 81 tuổi. Vợ cả Cụ Bạch Hưng Nghi"),
    (77, "20 tháng Mười Một", "Bà Hồ Thị Huyền", "Can 2 - Đời 6", "Rú Trước", "Thọ 78 tuổi (1935-2013). Vợ Liệt sỹ Bạch Hưng Xuân"),
    (78, "30 tháng Mười Một", "Cụ bà Nguyễn Thị Em", "Can 2 - Đời 5", "Núi Khe Xứ", "Thọ 76 tuổi (1890-1964). Vợ Cụ Tam (Chắt Mãi)"),
    (79, "04 tháng Mười Hai", "Cụ bà Ngô Thị Chuẩn", "Can 2 - Đời 4", "Thồ Lô xứ", "Vợ Cụ Bạch Hưng Xu"),
    (80, "06 tháng Giêng", "Bà Ngô Thị Thái", "Can 2 - Đời 7", "", "Mất 06/01/2016 Bính Thân (13/02/2016 DL). Vợ ông Bạch Hưng Uyển"),
    (81, "09 tháng Giêng", "Cụ bà Hồ Thị Tam", "Can 2 - Đời 6", "", "Thọ 90 tuổi (1932-2021). Mất 09/01 ÂL (Nhâm Thân). Vợ ông Bạch Hưng Bốn (Uyển), con gái cụ Cựu Học"),
    (81, "15 tháng Mười Hai", "Cụ bà Ngô Thị Chử", "Can 2 - Đời 5", "Núi Khe xứ", "Thọ 77 tuổi (1901-1978). Vợ Cụ Em (Con)"),
    (82, "30 tháng Mười Hai", "Cụ Bạch Hưng Nguyễn", "Can 3 - Đời 5", "Núi Khe xứ", "Con cả Cụ Truyện")
]

ws_gio = wb.create_sheet(title="Bảng Kỵ Giỗ")
gio_cols_def = [
    ("STT", 6, align_center),
    ("Ngày Âm lịch", 18, align_center),
    ("Người được giỗ", 24, align_left),
    ("Đời / Can", 18, align_center),
    ("Mộ phần / Nơi an táng", 24, align_left),
    ("Ghi chú kỵ giỗ (Theo văn bản)", 35, align_left),
]
ws_gio.views.sheetView[0].showGridLines = True
ws_gio.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(gio_cols_def))
c1 = ws_gio.cell(row=1, column=1, value="BẢNG TRA CỨU NGÀY KỴ GIỖ ÂM LỊCH TOÀN DÒNG HỌ (THEO VĂN BẢN GỐC)")
c1.font = font_title
c1.alignment = align_center

ws_gio.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(gio_cols_def))
c2 = ws_gio.cell(row=2, column=1, value="SẮP XẾP TUẦN TỰ THEO 12 THÁNG ÂM LỊCH TỪ THÁNG GIÊNG ĐẾN THÁNG CHẠP")
c2.font = font_subtitle
c2.alignment = align_center

for c_idx, (c_name, c_width, c_align) in enumerate(gio_cols_def, start=1):
    cell = ws_gio.cell(row=4, column=c_idx, value=c_name)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = thin_border
    col_letter = get_column_letter(c_idx)
    ws_gio.column_dimensions[col_letter].width = c_width

ws_gio.row_dimensions[1].height = 28
ws_gio.row_dimensions[2].height = 20
ws_gio.row_dimensions[4].height = 26

for r_idx, item in enumerate(gio_list, start=5):
    vals = [item[0], item[1], item[2], item[3], item[4], item[5]]
    is_zebra = (r_idx % 2 == 0)
    for c_idx, val in enumerate(vals, start=1):
        c = ws_gio.cell(row=r_idx, column=c_idx, value=val)
        c.font = font_bold if c_idx in [2, 3] else font_body
        c.border = thin_border
        c.alignment = gio_cols_def[c_idx-1][2]
        if is_zebra:
            c.fill = fill_zebra
    ws_gio.row_dimensions[r_idx].height = 22

wb.save(EXCEL_PATH)
print("Đã tạo file Excel thành công!")

# -------------------------------------------------------------
# 2. TẠO FILE WORD HOÀN CHỈNH
# -------------------------------------------------------------
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
r3 = p_cov.add_run("(PHÂN TÍCH CHUẨN XÁC DỰA 100% THEO VĂN BẢN GỐC)")
r3.font.name = "Times New Roman"; r3.font.size = Pt(12); r3.font.italic = True; r3.font.color.rgb = COLOR_SECONDARY

for _ in range(4): doc.add_paragraph()
p_loc = doc.add_paragraph()
p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_loc.add_run("LÀNG PHÚ ĐIỀN - XÃ HƯNG PHÚ - HUYỆN HƯNG NGUYÊN - TỈNH NGHỆ AN\nNĂM GIÁP NGỌ (2014) - SỐ HÓA HIỆU ĐÍNH NĂM 2026")
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
add_p("ĐẶC BIỆT LƯU Ý: Trong tập sách gốc năm 2014, một số trang sơ đồ vẽ hình khối (diagram) do người vẽ dồn cột, nối nét giản lược nên đã xảy ra sự sai lệch rất lớn về quan hệ cha con so với lời văn mô tả chi tiết. Toàn bộ bản số hóa này đã được hiệu chỉnh tuyệt đối trung thực dựa 100% VÀO VĂN BẢN CHỮ VIẾT (TEXT) của từng đời, bảo đảm tính xác thực gia phả đời đời.")

add_h1("VÀI NÉT VỀ LỊCH SỬ HỌ BẠCH & DANH NHÂN BẠCH LIÊU")
add_h2("1. Thủy tổ Trạng nguyên Bạch Liêu (1236 - 1315)")
add_p("Dòng họ Bạch Việt Nam tự hào có Cụ Thủy tổ là Trạng nguyên Bạch Liêu, sinh ngày 15 tháng Giêng năm Bính Thìn (1236) tại làng Yên Xá, huyện Chân Phúc (nay thuộc xã Mã Thành, huyện Yên Thành, tỉnh Nghệ An). Năm Bính Dần (1266) đời vua Trần Thánh Tông, ngài đỗ Trạng nguyên (Đệ nhất giáp tiến sĩ đệ nhất danh) khi tròn 31 tuổi.")
add_p("Ngài là Môn khách danh dự của Thượng phụ Chiêu Minh Đại vương Trần Quang Khải, có công lao to lớn hoạch định cơ mưu đánh giặc Nguyên - Mông và soạn thảo 'Biến pháp tam chương' cải cách quốc gia. Cụ tạ thế ngày 24 tháng Giêng năm Ất Mão (1315), thọ 80 tuổi.")
add_h2("2. Đền thờ họ Bạch tại Hưng Phú, Hưng Nguyên")
add_p("Đền thờ họ Bạch tại xã Hưng Phú đã được Nhà nước xếp hạng Di tích Lịch sử - Văn hóa cấp Tỉnh theo Quyết định số 2015/QĐ-UBND ngày 16/12/1993, được tôn tạo khang trang năm 2013.")

# Tổ Khởi phát & Đời 1, 2
add_h1("TỔ KHỞI PHÁT TRUNG CHI 2 & ĐỜI THỨ NHẤT, THỨ HAI")
add_h2("Tổ Khởi Phát Trung Chi 2: Cụ BẠCH HƯNG HÓA (1730 - 1802)")
add_p("Sinh năm Canh Tuất (1730), mất năm Nhâm Tuất (1802) thọ 73 tuổi. Giỗ ngày 14/04 ÂL. Mộ táng xứ Cồn Xanh. Ba cụ bà: Bùi Thị Hàng Nhất (giỗ 07/08 ÂL, mộ Cồn Xanh), Bùi Thị Hàng Hai (giỗ 28/01 ÂL, mộ Khe Dài), Nguyễn Thị Hàng Hai (giỗ 11/08 ÂL, mộ Khe Dài). Sinh 8 người con: Bạch Trần Tạo (TC1), Bạch Trần Điển (TC2), Bạch Trần Khơm (TC3), Bạch Khắc Chẩn, Bạch Hưng Đại, Bạch Thị Thao, Bạch Thị Đỏ...")

add_h2("Đời Thứ Nhất: Cụ BẠCH TRẦN TẠO (1772 - 1801)")
add_p("Khởi tổ Tiểu Chi 1. Sinh năm Nhâm Thìn (1772), mất năm Tân Dậu (1801) thọ 30 tuổi. Giỗ ngày 28/09 ÂL. Mộ đặt tại xứ Tuần Tượng. Vợ là Cụ bà Phạm Thị Hàng Nhất, thọ 76 tuổi, giỗ 08/10 ÂL, mộ Cồn Xanh. Sinh 2 con trai: Bạch Hưng Khâm, Bạch Hưng Lưu.")

add_h2("Đời Thứ Hai: Cụ BẠCH HƯNG KHÂM & Cụ BẠCH HƯNG LƯU")
add_p("1. Cụ Bạch Hưng Khâm: Con trưởng Cụ Tạo, mất năm 20 tuổi chưa lập gia đình. Giỗ ngày 28/05 ÂL, mộ Cồn Vải. Nhánh Cụ Lưu phụng tự.")
add_p("2. Cụ Bạch Hưng Lưu (húy Túc): Sinh năm Mậu Ngọ (1798), mất năm Quý Dậu (1873) thọ 76 tuổi. Giỗ ngày 29/02 ÂL, mộ xứ Thượng Phù. Vợ là Cụ bà Phạm Thị Biềng (giỗ 10/01 ÂL, mộ Thượng Phù). Sinh 4 con trai chia làm 4 Can: Cư (Can 1), Khâu (Can 2), Nghi (Can 3), Kế (Can 4); và 4 con gái: Thuật, Khoát, Khoan, Út.")

# CAN 1
add_h1("CAN THỨ NHẤT: TỔ CAN BẠCH HƯNG CƯ (ĐỜI 3 ĐẾN ĐỜI 8)")
add_h2("Đời Thứ 3: Cụ BẠCH HƯNG CƯ (húy Khuơ)")
add_p("Sinh năm Kỷ Mão (1819). Con trai cả Cụ Lưu. Giỗ ngày 09/10 ÂL. Mộ tại núi Khe xứ.")
add_p("- Vợ cả: Nguyễn Thị Máy (giỗ 08/02 ÂL, không thấy ghi lại con cái).")
add_p("- Vợ: Ngô Thị Hạng Nhị (giỗ 22/10 ÂL), sinh: Hoạch (chết nhỏ), Khương, Ba, Sáu, Hạnh (chết nhỏ), Thiết (húy Khải), Ý.")
add_p("- Vợ hai: Nguyễn Quý Nương (húy Vân, thọ 85, giỗ 28/03 ÂL), sinh: Tơn, Ngân, Lùa (chết nhỏ), Tiu.")

add_h2("Đời Thứ 4 và Đời Thứ 5 Can 1 (Theo văn bản trang 21 - 24)")
add_p("1. Bạch Hưng Thiết (húy Khải): Giỗ 01/03 ÂL, mộ Khe Xứ. Vợ Phan Thị Nghi. Sinh con trai độc nhất là Bạch Hưng Khớt (có vợ không con, mất năm 46 tuổi).")
add_p("2. Bạch Hưng Ý: Con thứ 3 bà Ngô Thị Hạng Nhị. Đi lính triều Nguyễn, chết đuối ở sông Hiếu, Nghĩa Đàn. Giỗ 01/10 ÂL, chiêu hồn mộ núi Khe xứ.")
add_p("3. Bạch Hưng Tơn: Thọ 60 tuổi, giỗ 12/06 ÂL. Vợ Phạm Thị Phòng (giỗ 07/05 ÂL). Sinh 3 con trai: Bạch Hưng Em (Bâu), Bạch Hưng Chế (Tam), Bạch Hưng Cận (Tam); và con gái Bạch Thị Tiu.")
add_p("4. Bạch Hưng Ngân: Giỗ 30/06 ÂL, mộ Rú Sau (cây đa Bò Mẹp). Vợ cả Hồ Thị Niêm (giỗ 15/07 ÂL) sinh 2 con trai: Bạch Hưng Khôn, Bạch Hưng Hai. Vợ hai Nguyễn Thị Em (húy Tảo, giỗ 17/09 ÂL) sinh 2 gái: Tứ, Ngụ.")
add_p("5. Cụ Bạch Hưng Chế (Tam) con Cụ Tơn: Thọ 74 tuổi, giỗ 19/01 ÂL, mộ núi Khe xứ. Vợ Nguyễn Thị Tứ (1900-1987, thọ 88, giỗ 11/10 ÂL). Cụ bà sinh được 5 người con trai là 5 anh em ruột: Bạch Hưng Cháu (1), Bạch Hưng Nhã (2), Bạch Hưng Mận (3), Bạch Hưng Mai (4), Bạch Hưng Nhỏ (5); và con gái Bạch Thị Mơ.")

add_h2("Đời Thứ 6 và Đời Thứ 7 Can 1 (Theo văn bản trang 25 - 32)")
add_p("1. Bạch Hưng Cháu (sinh 1923, mất 2014 Giáp Ngọ, giỗ 30/06 ÂL, mộ Núi Khe Xứ, vợ Nguyễn Thị Giá sinh 1923). Sinh 4 trai: Bạch Hưng Tùng (sinh 1951, vợ Lam Trà, con: Hương Giang, Kiên); Bạch Hưng Hùng (sinh 1956, vợ Hồ Thị Sơn, con: Thùy Dương, Thủy Tiên, Trà Hoa Hạ, Đạt); Bạch Hưng Dũng (sinh 1959, Thượng tá CSGT Công an Nghệ An, vợ Phạm Thị Mười, con: Hương Trà, Hương Thảo); Bạch Hưng Thất (sinh 1964, vợ Hồng Nhạn, con: Bạch Hưng An 1997, Hải Lê 1999); và 3 gái: Thành, Phùng, Tâm.")
add_p("2. Bạch Hưng Nhạ (1923 - 1998, thọ 76, giỗ 23/03 ÂL, vợ Nguyễn Thị Chắt 1926-2005 thọ 80). Sinh 2 trai: Bạch Quốc Việt (sinh 1958, vợ Phạm Thị Tâm, con: Tuấn Dũng 1984, Hồng Duyên, Mỹ Lệ); Bạch Hưng Nam (sinh 1960, Chủ tịch xã Nghĩa Đức, vợ Nguyễn Thị Lan, con: Bảo Trung 1986, Khánh Chi, Thùy); và 7 con gái: Hòa, Nguyệt, Bảo, Toàn, Thảo, Hiền, Lương.")
add_p("3. Bạch Hưng Mận (1935 - 2002, thọ 68, giỗ 10/02 ÂL, vợ Nguyễn Thị Loan 1935-1972). Sinh 4 trai: Bạch Hưng Lập (1958-1968, mất nhỏ); Bạch Hưng Thanh (sinh 1961, vợ Hoa Thị Huyền, con: Hoàn, Thương, Hảo, Hạnh, Phúc); Bạch Hưng Diệu (sinh 1968, vợ Hồ Thị Hương, con: Tùng Dương 1996, Quang Sinh 2001); Bạch Hưng Lợi (sinh 1972, vợ Trần Thị Tư, con: Bạch Hưng Thắng 2001); và 2 gái: Minh, Oanh.")
add_p("4. Bạch Hưng Mai (sinh 1939, vợ Hồ Thị Tình sinh 1944). Sinh 1 trai: Bạch Hưng Hải (sinh 1981, vợ Phan Thị Hoa, con: Quốc Quân 2005, Quyên 2012); và 3 gái: Oanh, Ân, Yến.")
add_p("5. Bạch Hưng Nhỏ (Bạch Quốc Hưng, sinh 1945, bộ đội chống Mỹ, thợ cơ khí tại Hải Dương, vợ Đào Thị Nữ Trang sinh 1951). Sinh 2 trai: Bạch Thái Nguyên (sinh 1974, lái xe, vợ Trần Thị Anh, con: Khánh An 2010); Bạch Quang Ngọc (sinh 1979, Trưởng phòng vật tư Cty Thép Hoàng Long, vợ Bùi Thị Phương, con: Đức Thịnh 2013); và 2 gái: Hoàng Giang (1976), Thu Hà (1979).")

# CAN 2
add_h1("CAN THỨ HAI: TỔ CAN BẠCH HƯNG KHÂU (ĐỜI 3 ĐẾN ĐỜI 8)")
add_h2("Đời Thứ 3 và Đời Thứ 4 Can 2 (Theo văn bản trang 35 - 38)")
add_p("Tổ Can: Cụ Bạch Hưng Khâu (húy Hoài, 1821 - 1897, thọ 77, giỗ 02/08 ÂL, mộ Núi Khe xứ). Vợ cả Bùi Thị Muồng (giỗ 07/08 ÂL, không con). Vợ hai Phạm Thị Hoạt (thọ 62, giỗ 28/08 ÂL, mộ Cồn Xanh) sinh 5 trai 1 gái:")
add_p("1. Bạch Hưng Thấu: Mất khi chưa có gia thất, giỗ 05/04 ÂL, mộ Kẹt Đá xứ.")
add_p("2. Bạch Hưng Thưu: Giỗ 25/10 ÂL, mộ Thồ Lô xứ. Vợ Hồ Thị Út (giỗ 19/03 ÂL). Sinh con trai độc nhất là Bạch Sĩ Xum (Cố Hòe, giỗ 11/08 ÂL, vợ Hồ Thị Nhiên giỗ 15/03 ÂL) và con gái Bạch Thị Đĩ (chết nhỏ).")
add_p("3. Bạch Hưng Xu (húy Mâu, sinh 1854 Giáp Dần, giỗ 01/06 ÂL, mộ Thồ Lô xứ, vợ Ngô Thị Chuẩn giỗ 04/01 ÂL). Sinh 4 trai: Tam (Cố Chắt Mãi), Hợi, Yêm, Em (mất chưa vợ, giỗ 06/08 ÂL); và 3 gái: Chắt, Tứ, Em.")
add_p("4. Bạch Hưng Phù (húy Tư Lang, 1862 - 1912, thọ 51, giỗ 28/06 ÂL, mộ Khe Xứ): Văn bản ghi rõ: ÔNG KHÔNG LẬP GIA THẤT, MẤT KHI 51 TUỔI.")
add_p("5. Bạch Hưng Câu (1864 - 1894, thọ 31, giỗ 28/09 ÂL, mộ Núi Khe Xứ, vợ Ngô Thị Tư giỗ 24/09 ÂL). Sinh 3 con trai: Bạch Hưng Phòng (mất chưa vợ, giỗ 19/05 ÂL); Bạch Hưng Dĩ (thọ 61, giỗ 17/03 ÂL, vợ Nguyễn Thị Lù giỗ 09/05 ÂL); Bạch Hưng Em (húy Con, 1900 - 1981 thọ 82, giỗ 28/03 ÂL, vợ Ngô Thị Chử 1901-1978 thọ 77 giỗ 15/12 ÂL); và con gái Bạch Thị Đị.")

add_h2("Đời Thứ 5 và Đời Thứ 6 Can 2 (Theo văn bản trang 39 - 50)")
add_p("1. Cụ Bạch Hưng Tam (Cố Chắt Mãi, 1891 - 1950 thọ 60, giỗ 22/02 ÂL, con Cụ Xu). Tuần Kiểm của làng. Vợ Nguyễn Thị Em (1890 - 1964 thọ 76, giỗ 30/11 ÂL). Sinh con trai độc nhất là Liệt sỹ Bạch Hưng Đình và 5 gái (Tâm, Cháu, Nhỏ, Tìu, Ca).")
add_p("2. Liệt sỹ Bạch Hưng Đình (1925 - 1968, Chủ tịch xã Hưng Phú, hy sinh 27/08/1968, giỗ 03/07 ÂL). Vợ đầu Hồ Thị Hảo (1922-1951, giỗ 30/01 ÂL) sinh 2 trai: Bạch Hưng Thái (sinh 1945), Bạch Hưng Phú (sinh 1949). Vợ sau Nguyễn Thị Quế (sinh 1925) sinh 3 trai: Bạch Hưng Dương (sinh 1954), Bạch Hưng Tuyên (sinh 1958), Bạch Hưng Trung (sinh 1966); và con gái Bạch Thị Dung (sinh 1962).")
add_p("3. Nhánh Cụ Bạch Hưng Hợi (con Cụ Xu): Sinh 4 con trai: Bạch Hưng Nậy (1918-1993, sinh: Vinh, Quang, Trung); Bạch Hưng Tứ (1923-1999, sinh: Trường, Thường); Bạch Hưng Niêm (1930-2011, sinh: Thông, Tiến, Ty); Bạch Hưng Công (sinh 1938, công nhân NT 3.2, sinh: Tâm, Long, Phương, Kiều).")
add_p("4. Nhánh Cụ Bạch Hưng Dĩ (con Cụ Câu): Sinh 3 con trai: Bạch Hưng Nậy (sinh 2 gái Tuyết, Tý); Bạch Hưng Cường (1920 - 2016 thọ 97, Cán bộ Tiền khởi nghĩa, sinh 1 con trai duy nhất là Bạch Hưng Huy và 7 con gái); Bạch Hưng Bốn / Uyển (1930 - 1993 thọ 64, Thẩm phán TAND huyện, sinh 6 con trai: Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn và 3 con gái Nga, Thành, Công).")
add_p("5. Nhánh Cụ Bạch Hưng Em (húy Con, con Cụ Câu): Sinh 2 con trai: Bạch Hưng Lực (1931 - 2009 thọ 79, thương binh Điện Biên Phủ, sinh 4 con trai: Mậu, Phượng, Trường, Lường và 2 gái Chắt, Châu); Liệt sỹ Bạch Hưng Xuân (1932 - 1966 hy sinh tại Trường Sơn, sinh 2 trai: Bạch Hưng Cầm, Bạch Hưng Quyền và con gái Thư).")

add_h2("Đời Thứ 7 Can 2 Tiêu Biểu (Theo văn bản trang 51 - 61)")
add_p("- Bạch Hưng Thái (sinh 1945, con cả Liệt sỹ Đình): Thương binh chống Mỹ, con: Bạch Hưng Thư (sinh 1978).")
add_p("- Bạch Hưng Phú (sinh 1949, con thứ 2 Liệt sỹ Đình): Nhà giáo, phụng tự Tổ Can 4 Cụ Kế. Con: Bạch Hồng Y (1979), Bạch Ly Na (1983 - Thạc sĩ, BS Viện Nhi TW).")
add_p("- Bạch Hưng Dương (sinh 1954, con thứ 3 Liệt sỹ Đình): Đại tá Công an Nghệ An. Con: Hồng Hải (1987), Hoàn (1992).")
add_p("- Bạch Hưng Tuyên (sinh 1958, con thứ 4 Liệt sỹ Đình): GĐ Cty Thủy nông Nam Nghệ An. Con: Tân (1987), Thùy Linh.")
add_p("- Bạch Hưng Trung (sinh 1966, con út Liệt sỹ Đình): Trưởng trạm Ba Ra Nam Đàn. Con: Ninh (1999), Huyền An.")
add_p("- Bạch Hưng Huy (sinh 1952, con trai độc nhất Cụ Cường): GV cấp 3 tại Đức Trọng, Lâm Đồng. Con: Hường, Hà, Quỳnh.")
add_p("- 6 anh em con Cụ Bạch Hưng Bốn (Uyển): Thượng tá Bạch Hưng Uyển (sinh 1952); Đại tá CA Bạch Hưng Triển (sinh 1956); Thượng úy Bạch Hưng Lịch (sinh 1961); Đại tá PGS Bạch Hưng Thịnh (sinh 1963); Bạch Hưng Vĩnh (sinh 1971); Bạch Hưng Viễn (sinh 1975).")
add_p("- 4 anh em con Cụ Bạch Hưng Lực: Thượng úy Bạch Hưng Mậu (sinh 1958); Bạch Hưng Phượng (sinh 1964); Bạch Hưng Trường (sinh 1969); Bạch Hưng Lường (sinh 1972).")
add_p("- 2 anh em con Liệt sỹ Bạch Hưng Xuân: Kỹ sư Bạch Hưng Cầm (sinh 1962, TP Cơ khí Vinh, con: Bình Minh đạt HC Đồng Olympic Tiếng Anh 2015); Bạch Hưng Quyền (sinh 1964, Làng SOS Vinh, con: Khang, Liên).")

# CAN 3 & CAN 4
add_h1("CAN THỨ BA & CAN THỨ TƯ (ĐỜI 3 ĐẾN ĐỜI 8)")
add_h2("Can 3: Tổ Can Bạch Hưng Nghi (Thủ chỉ Phú Điền)")
add_p("Cụ Bạch Hưng Nghi (1829, thọ 89, giỗ 10/04 ÂL, mộ Núi Khe xứ). Sinh Cụ Bạch Hưng Truyện (23 tuổi làm Lý trưởng, 28 tuổi khẳng khái bị hại, giỗ 11/07 ÂL). Cụ Truyện sinh 2 con trai: Bạch Hưng Nguyễn (giỗ 30/12 ÂL) và Cai xã Bạch Hưng Khuyến (thọ 66, giỗ 06/09 ÂL).")
add_p("- Nhánh Cụ Nguyễn: Cụ Bạch Hưng Hai (Kiểm Nhưng, thọ 54, giỗ 20/08 ÂL) sinh Bạch Hưng Thành (1936-1985, Bí thư đoàn, Cán bộ GT 470, HC Kháng chiến Hạng Nhất, sinh 5 trai: Chiến, Sỹ, Hương, Thủy, Triều và 4 gái); Cụ Bạch Hưng Tư (sinh 1906) sinh Thương binh nặng 1/4 Bạch Hưng Liên (sinh 1943, lái xe Đường 9 Nam Lào hỏng 2 mắt, sinh: Thắng, Đức, Lài); Cụ Bạch Hưng Niêm (mất chưa vợ); Cụ Bạch Hưng Sáu (mất tuổi 35).")
add_p("- Nhánh Cụ Khuyến: Cụ Bạch Hưng Hiến (Đầu phiên, 1906-1986, thọ 82, giỗ 15/02 ÂL) sinh Thương binh 2/4 Bạch Hưng Văn (sinh 1945, pháo cao xạ cụt bàn tay trái, sinh: Giang, Tình, Yến, Oanh, Nhung) và Cử nhân Kinh tế Bạch Hưng Minh (1948-1993 thọ 47, KT trưởng Cty Vật tư NN, sinh: Sơn, Hải - Thượng úy QK4, Hiền); Cụ Bạch Hưng Hưng (mất năm 14 tuổi).")

add_h2("Can 4: Tổ Can Bạch Hưng Kế (Húy Cơ)")
add_p("Sinh năm Giáp Dần (1854). Con trai thứ 4 Cụ Lưu. Vợ Ngô Thị Tam (húy Chư). Con trai Thế mất năm 13 tuổi, 2 gái Mây và Út. Lăng mộ tại xứ Thồ Lô. Cụ vô tự, chi họ phụng tự. Trước đây ông Bạch Hưng Đình đảm nhận; nay con thứ 2 là ông Bạch Hưng Phú tiếp nối thờ tự chu toàn tại gia đường.")

doc.save(DOCX_PATH)
print("Đã tạo file Word thành công!")
