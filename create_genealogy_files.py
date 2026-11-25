# -*- coding: utf-8 -*-
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUTPUT_DIR = r"d:\python\gia pha ho Bach"
EXCEL_PATH = os.path.join(OUTPUT_DIR, "Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx")
DOCX_PATH = os.path.join(OUTPUT_DIR, "Gia_Pha_Ho_Bach.docx")

print("Bắt đầu xử lý dữ liệu gia phả...")

# -----------------------------------------------------------------------------
# DỮ LIỆU NHÂN ĐINH ĐẦY ĐỦ (DATABASE)
# -----------------------------------------------------------------------------
records = [
    # KHỞI PHÁT & TIỀN ĐẠI
    {
        "stt": 1, "ma": "KP-01", "can": "Trung Chi 2", "doi": "Khởi phát",
        "ho_ten": "Bạch Hưng Hóa", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1730 (Canh Tuất)", "nam_mat": "1802 (Nhâm Tuất) - Thọ 73",
        "ngay_gio": "14/04 ÂL", "cha": "", "me": "",
        "phoi_ngau": "Bùi Thị Hàng Nhất; Bùi Thị Hàng Hai; Nguyễn Thị Hàng Hai",
        "mo_phan": "Xứ Cồn Xanh",
        "ghi_chu": "Tổ Khởi phát Trung Chi 2 họ Bạch, sinh 8 người con (Bạch Trần Tạo, Bạch Trần Điển, Bạch Trần Khơm, Bạch Khắc Chẩn, Bạch Hưng Đại, Bạch Thị Thao, Bạch Thị Đỏ...)"
    },
    {
        "stt": 2, "ma": "TC1-D1-01", "can": "Tiểu Chi 1", "doi": "Đời 1",
        "ho_ten": "Bạch Trần Tạo", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1772 (Nhâm Thìn)", "nam_mat": "1801 (Tân Dậu) - Thọ 30",
        "ngay_gio": "28/09 ÂL", "cha": "Bạch Hưng Hóa", "me": "Bùi Thị Hàng Nhất",
        "phoi_ngau": "Phạm Thị Hàng Nhất (Thọ 76, giỗ 08/10 ÂL, mộ Cồn Xanh)",
        "mo_phan": "Xứ Tuần Tượng",
        "ghi_chu": "Khởi tổ Tiểu Chi 1 thuộc Trung Chi 2 họ Bạch. Sinh 2 con trai: Bạch Hưng Khâm, Bạch Hưng Lưu."
    },
    {
        "stt": 3, "ma": "TC1-D2-01", "can": "Tiểu Chi 1", "doi": "Đời 2",
        "ho_ten": "Bạch Hưng Khâm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất năm 20 tuổi",
        "ngay_gio": "28/05 ÂL", "cha": "Bạch Trần Tạo", "me": "Phạm Thị Hàng Nhất",
        "phoi_ngau": "Chưa lập gia thất",
        "mo_phan": "Xứ Cồn Vải",
        "ghi_chu": "Con trưởng Cụ Bạch Trần Tạo, mất sớm. Nhánh Cụ Lưu phụng tự."
    },
    {
        "stt": 4, "ma": "TC1-D2-02", "can": "Tiểu Chi 1", "doi": "Đời 2",
        "ho_ten": "Bạch Hưng Lưu", "huy": "Húy Túc", "gioi_tinh": "Nam",
        "nam_sinh": "1798 (Mậu Ngọ)", "nam_mat": "1873 (Quý Dậu) - Thọ 76",
        "ngay_gio": "29/02 ÂL", "cha": "Bạch Trần Tạo", "me": "Phạm Thị Hàng Nhất",
        "phoi_ngau": "Phạm Thị Biềng (Giỗ 10/01 ÂL, mộ xứ Thượng Phù)",
        "mo_phan": "Xứ Thượng Phù",
        "ghi_chu": "Con thứ Cụ Tạo. Sinh 4 con trai phân thành 4 Can (Cư, Khâu, Nghi, Kế) và 4 con gái (Thuật, Khoát, Khoan, Út)."
    },

    # =========================================================================
    # CAN 1: TỔ CỤ BẠCH HƯNG CƯ
    # =========================================================================
    {
        "stt": 5, "ma": "C1-D3-01", "can": "Can 1", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Cư", "huy": "Húy Khuơ", "gioi_tinh": "Nam",
        "nam_sinh": "1819 (Kỷ Mão)", "nam_mat": "",
        "ngay_gio": "09/10 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Nguyễn Thị Máy (giỗ 08/02 ÂL); Bà 2: Ngô Thị Hạng Nhị (giỗ 22/10 ÂL); Bà 3: Nguyễn Quý Nương (hiệu Từ Thuần, húy Vân, thọ 85, giỗ 28/03 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 1. Con trai trưởng Cụ Lưu. Sinh 4 gái (Hoạch, Khương, Ba, Sáu) và 7 trai (Hạnh, Thiết, Ý, Tơn, Ngân, Lùa, Tiu)."
    },
    {
        "stt": 6, "ma": "C1-D4-01", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thiết", "huy": "Húy Khuyến", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cư", "me": "",
        "phoi_ngau": "Nguyễn Thị Thao",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai Cụ Cư. Sinh 2 con trai: Bạch Hưng Khớt, Bạch Hưng Em (Bâu)."
    },
    {
        "stt": 7, "ma": "C1-D4-02", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Ý", "huy": "Húy Giám", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cư", "me": "",
        "phoi_ngau": "Nguyễn Thị Tình",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai Cụ Cư. Sinh con trai: Bạch Hưng Chế (Tam)."
    },
    {
        "stt": 8, "ma": "C1-D4-03", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Tơn", "huy": "Húy Cửu", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cư", "me": "",
        "phoi_ngau": "Đậu Thị Nhị",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai Cụ Cư. Sinh con trai: Bạch Hưng Cận."
    },
    {
        "stt": 9, "ma": "C1-D4-04", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Ngân", "huy": "Húy Khoản", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cư", "me": "",
        "phoi_ngau": "Phạm Thị Kiều",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai Cụ Cư. Sinh 2 con trai: Bạch Hưng Khôn, Bạch Hưng Hai (Hiểm)."
    },
    {
        "stt": 10, "ma": "C1-D5-01", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Khớt", "huy": "Húy Nhược", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Thiết", "me": "Nguyễn Thị Thao",
        "phoi_ngau": "Nguyễn Thị Duyên",
        "mo_phan": "",
        "ghi_chu": "Sinh 3 con trai: Bạch Hưng Chắt, Bạch Hưng Chính, Bạch Hưng Mô."
    },
    {
        "stt": 11, "ma": "C1-D5-02", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Tên gọi: Bâu", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Thiết", "me": "Nguyễn Thị Thao",
        "phoi_ngau": "Phạm Thị Mân",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Cháu."
    },
    {
        "stt": 12, "ma": "C1-D5-03", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Chế", "huy": "Tên gọi: Tam", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Ý", "me": "Nguyễn Thị Tình",
        "phoi_ngau": "Trần Thị Thoan",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Nhã."
    },
    {
        "stt": 13, "ma": "C1-D5-04", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Cận", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tơn", "me": "Đậu Thị Nhị",
        "phoi_ngau": "Nguyễn Thị Lợi",
        "mo_phan": "",
        "ghi_chu": "Sinh 2 con trai: Bạch Hưng Mận, Bạch Hưng Mai."
    },
    {
        "stt": 14, "ma": "C1-D5-05", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Hai", "huy": "Tên gọi: Hiểm", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Ngân", "me": "Phạm Thị Kiều",
        "phoi_ngau": "Phan Thị Tý",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Nhỏ."
    },
    # Đời 6 Can 1
    {
        "stt": 15, "ma": "C1-D6-01", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Chắt", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Khớt", "me": "Nguyễn Thị Duyên",
        "phoi_ngau": "Lê Thị Phượng",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Tùng, Bạch Hưng Hùng, Bạch Thị Oanh."
    },
    {
        "stt": 16, "ma": "C1-D6-02", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Chính", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Khớt", "me": "Nguyễn Thị Duyên",
        "phoi_ngau": "Phạm Thị Vân",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Dũng, Bạch Thị Nga."
    },
    {
        "stt": 17, "ma": "C1-D6-03", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mô", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Khớt", "me": "Nguyễn Thị Duyên",
        "phoi_ngau": "Nguyễn Thị Hương",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Thất, Bạch Quốc Việt, Bạch Hưng Nam."
    },
    {
        "stt": 18, "ma": "C1-D6-04", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Cháu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Bâu)", "me": "Phạm Thị Mân",
        "phoi_ngau": "Võ Thị Kim",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Lập, Bạch Thị Huyền."
    },
    {
        "stt": 19, "ma": "C1-D6-05", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nhã", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chế (Tam)", "me": "Trần Thị Thoan",
        "phoi_ngau": "Trần Thị Mai",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Thanh, Bạch Hưng Diệu."
    },
    {
        "stt": 20, "ma": "C1-D6-06", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mận", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cận", "me": "Nguyễn Thị Lợi",
        "phoi_ngau": "Nguyễn Thị Hoa",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Lợi, Bạch Thị Sen."
    },
    {
        "stt": 21, "ma": "C1-D6-07", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mai", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cận", "me": "Nguyễn Thị Lợi",
        "phoi_ngau": "Ngô Thị Quyên",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Hải, Bạch Thái Nguyên."
    },
    {
        "stt": 22, "ma": "C1-D6-08", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nhỏ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hai (Hiểm)", "me": "Phan Thị Tý",
        "phoi_ngau": "Nguyễn Thị Loan",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Quang Ngọc, Bạch Thị Ngọc Anh."
    },
    # Đời 7 & 8 Can 1
    {
        "stt": 23, "ma": "C1-D7-01", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tùng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chắt", "me": "Lê Thị Phượng",
        "phoi_ngau": "Nguyễn Thị Huệ",
        "mo_phan": "",
        "ghi_chu": "Ban biên tập Gia phả năm 2014. Sinh: Bạch Hưng Cường, Bạch Hưng Tuấn."
    },
    {
        "stt": 24, "ma": "C1-D7-02", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Hùng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chắt", "me": "Lê Thị Phượng",
        "phoi_ngau": "Lê Thị Mai",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Long."
    },
    {
        "stt": 25, "ma": "C1-D7-03", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Dũng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chính", "me": "Phạm Thị Vân",
        "phoi_ngau": "Phạm Thị Sen",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Bình."
    },
    {
        "stt": 26, "ma": "C1-D7-04", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thất", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mô", "me": "Nguyễn Thị Hương",
        "phoi_ngau": "Hoàng Thị Hà",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Phú, Bạch Hưng Quý."
    },
    {
        "stt": 27, "ma": "C1-D7-05", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Quốc Việt", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mô", "me": "Nguyễn Thị Hương",
        "phoi_ngau": "Nguyễn Thị Thanh",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng An."
    },
    {
        "stt": 28, "ma": "C1-D7-06", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Nam", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mô", "me": "Nguyễn Thị Hương",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai thứ 3 ông Mô."
    },
    {
        "stt": 29, "ma": "C1-D7-07", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lập", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cháu", "me": "Võ Thị Kim",
        "phoi_ngau": "Đậu Thị Lý",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Phúc."
    },
    {
        "stt": 30, "ma": "C1-D7-08", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thanh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhã", "me": "Trần Thị Mai",
        "phoi_ngau": "Võ Thị Trang",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Khang."
    },
    {
        "stt": 31, "ma": "C1-D7-09", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Diệu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhã", "me": "Trần Thị Mai",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con thứ 2 ông Nhã."
    },
    {
        "stt": 32, "ma": "C1-D7-10", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lợi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mận", "me": "Nguyễn Thị Hoa",
        "phoi_ngau": "Lê Thị Hằng",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Đạt."
    },
    {
        "stt": 33, "ma": "C1-D7-11", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Hải", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mai", "me": "Ngô Thị Quyên",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con cả ông Mai."
    },
    {
        "stt": 34, "ma": "C1-D7-12", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Thái Nguyên", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mai", "me": "Ngô Thị Quyên",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con thứ 2 ông Mai."
    },
    {
        "stt": 35, "ma": "C1-D7-13", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Quang Ngọc", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhỏ", "me": "Nguyễn Thị Loan",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai ông Nhỏ."
    },

    # =========================================================================
    # CAN 2: TỔ CỤ BẠCH HƯNG KHÂU
    # =========================================================================
    {
        "stt": 36, "ma": "C2-D3-01", "can": "Can 2", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Khâu", "huy": "Húy Hoài", "gioi_tinh": "Nam",
        "nam_sinh": "1821 (Tân Tỵ)", "nam_mat": "1897 (Đinh Dậu) - Thọ 77",
        "ngay_gio": "02/08 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Bùi Thị Muồng (giỗ 07/08 ÂL, mộ Rộc Mâm); Bà 2: Phạm Thị Hoạt (thọ 62, giỗ 28/08 ÂL, mộ Cồn Xanh)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 2. Con trai thứ 2 Cụ Lưu. Sinh 5 trai (Thấu, Thưu, Xu, Phù, Câu) và 1 gái (Bạch Thị Hai)."
    },
    {
        "stt": 37, "ma": "C2-D4-01", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thấu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi còn nhỏ",
        "ngay_gio": "20/01 ÂL", "cha": "Bạch Hưng Khâu", "me": "",
        "phoi_ngau": "Chưa lập gia thất", "mo_phan": "Xứ Đồng Mồ",
        "ghi_chu": "Con trai trưởng Cụ Khâu."
    },
    {
        "stt": 38, "ma": "C2-D4-02", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thưu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "01/11 ÂL", "cha": "Bạch Hưng Khâu", "me": "",
        "phoi_ngau": "Vô tự", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 2 Cụ Khâu."
    },
    {
        "stt": 39, "ma": "C2-D4-03", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Xu", "huy": "Húy Mâu", "gioi_tinh": "Nam",
        "nam_sinh": "1854 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Khâu", "me": "",
        "phoi_ngau": "Phan Thị Tác (Giỗ 22/10 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 3 Cụ Khâu. Sinh 5 con trai: Xum (Hòe - mất sớm), Tam (Chắt Mãi), Hợi, Yêm, Em (Bốn)."
    },
    {
        "stt": 40, "ma": "C2-D4-04", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Phù", "huy": "Húy Tư Lang", "gioi_tinh": "Nam",
        "nam_sinh": "1862 (Nhâm Tuất)", "nam_mat": "1912 (Nhâm Tý) - Thọ 51",
        "ngay_gio": "18/01 ÂL", "cha": "Bạch Hưng Khâu", "me": "",
        "phoi_ngau": "Đậu Thị Diệu (Thọ 66, giỗ 14/10 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 4 Cụ Khâu. Sinh 3 con trai: Bạch Hưng Phòng, Bạch Hưng Dị, Bạch Hưng Em (Con)."
    },
    {
        "stt": 41, "ma": "C2-D4-05", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Câu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1864 (Giáp Tý)", "nam_mat": "1894 (Giáp Ngọ) - Thọ 31",
        "ngay_gio": "21/04 ÂL", "cha": "Bạch Hưng Khâu", "me": "",
        "phoi_ngau": "Nguyễn Thị Tình",
        "mo_phan": "",
        "ghi_chu": "Con trai thứ 5 Cụ Khâu."
    },
    # Đời 5 Can 2
    {
        "stt": 42, "ma": "C2-D5-01", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Tam", "huy": "Tên gọi: Chắt Mãi", "gioi_tinh": "Nam",
        "nam_sinh": "1891 (Tân Mão)", "nam_mat": "1950 (Canh Dần) - Thọ 60",
        "ngay_gio": "21/04 ÂL", "cha": "Bạch Hưng Xu", "me": "Phan Thị Tác",
        "phoi_ngau": "Nguyễn Thị Trúc (Thọ 74, giỗ 18/06 ÂL)",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai độc nhất: Liệt sỹ Bạch Hưng Đình."
    },
    {
        "stt": 43, "ma": "C2-D5-02", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Hợi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xu", "me": "Phan Thị Tác",
        "phoi_ngau": "Hoàng Thị Tý",
        "mo_phan": "",
        "ghi_chu": "Sinh 4 con trai: Bạch Hưng Nậy (Xuân), Bạch Hưng Tứ, Bạch Hưng Niêm, Bạch Hưng Công."
    },
    {
        "stt": 44, "ma": "C2-D5-03", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Yêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xu", "me": "Phan Thị Tác",
        "phoi_ngau": "Đinh Thị Cháu",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Nậy (Yêm)."
    },
    {
        "stt": 45, "ma": "C2-D5-04", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Tên gọi: Bốn", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xu", "me": "Phan Thị Tác",
        "phoi_ngau": "Đậu Thị Tám",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Vỹ."
    },
    {
        "stt": 46, "ma": "C2-D5-05", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Phòng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Phù", "me": "Đậu Thị Diệu",
        "phoi_ngau": "Nguyễn Thị Nuôi",
        "mo_phan": "",
        "ghi_chu": "Con cả Cụ Phù."
    },
    {
        "stt": 47, "ma": "C2-D5-06", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Dị", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Phù", "me": "Đậu Thị Diệu",
        "phoi_ngau": "Đậu Thị Mai",
        "mo_phan": "",
        "ghi_chu": "Sinh con trai: Bạch Hưng Nậy (Dị)."
    },
    {
        "stt": 48, "ma": "C2-D5-07", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Tên gọi: Con", "gioi_tinh": "Nam",
        "nam_sinh": "1900 (Canh Tý)", "nam_mat": "1981 (Tân Dậu) - Thọ 82",
        "ngay_gio": "20/01 ÂL", "cha": "Bạch Hưng Phù", "me": "Đậu Thị Diệu",
        "phoi_ngau": "Phạm Thị Thiện (Thọ 67, giỗ 19/03 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Sinh 4 con trai: Bạch Hưng Cường, Bạch Hưng Bốn (Uyển), Bạch Hưng Lực, Liệt sỹ Bạch Hưng Xuân."
    },
    # Đời 6 Can 2
    {
        "stt": 49, "ma": "C2-D6-01", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Đình", "huy": "Liệt sỹ", "gioi_tinh": "Nam",
        "nam_sinh": "1925 (Ất Sửu)", "nam_mat": "1968 (Mậu Thân) - Hy sinh",
        "ngay_gio": "12/04 ÂL (hoặc theo ngày báo tử)", "cha": "Bạch Hưng Tam", "me": "Nguyễn Thị Trúc",
        "phoi_ngau": "Nguyễn Thị Mậu (1923 - 2012, thọ 89)",
        "mo_phan": "Nghĩa trang Liệt sỹ",
        "ghi_chu": "Liệt sỹ chống Mỹ, hy sinh tại chiến trường Quảng Trị. Sinh: Bạch Hưng Thái, Bạch Hưng Phú, Bạch Hưng Dương, Bạch Thị Phượng, Bạch Thị Lương."
    },
    {
        "stt": 50, "ma": "C2-D6-02", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Tên gọi: Xuân", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hợi", "me": "Hoàng Thị Tý",
        "phoi_ngau": "Nguyễn Thị Sen",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Tuyên, Bạch Hưng Trung."
    },
    {
        "stt": 51, "ma": "C2-D6-03", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Tứ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hợi", "me": "Hoàng Thị Tý",
        "phoi_ngau": "Đậu Thị Lộc",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Vinh, Bạch Thị Lan."
    },
    {
        "stt": 52, "ma": "C2-D6-04", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Niêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hợi", "me": "Hoàng Thị Tý",
        "phoi_ngau": "Lê Thị Mai",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Quang."
    },
    {
        "stt": 53, "ma": "C2-D6-05", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Công", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hợi", "me": "Hoàng Thị Tý",
        "phoi_ngau": "Hoàng Thị Thanh",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Trung."
    },
    {
        "stt": 54, "ma": "C2-D6-06", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Tên gọi: Yêm", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Yêm", "me": "Đinh Thị Cháu",
        "phoi_ngau": "Phan Thị Lài",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Trường, Bạch Hưng Thường."
    },
    {
        "stt": 55, "ma": "C2-D6-07", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Vỹ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Bốn)", "me": "Đậu Thị Tám",
        "phoi_ngau": "Trần Thị Tuyết",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Thông, Bạch Hưng Tiến."
    },
    {
        "stt": 56, "ma": "C2-D6-08", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Tên gọi: Dị", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Dị", "me": "Đậu Thị Mai",
        "phoi_ngau": "Phạm Thị Lan",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Ty, Bạch Hưng Tâm."
    },
    {
        "stt": 57, "ma": "C2-D6-09", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Cường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1920 (Canh Thân)", "nam_mat": "2016 (Bính Thân) - Thọ 97",
        "ngay_gio": "24/05 ÂL", "cha": "Bạch Hưng Em (Con)", "me": "Phạm Thị Thiện",
        "phoi_ngau": "Hoàng Thị Dung (sinh 1921)",
        "mo_phan": "Đức Trọng, Lâm Đồng",
        "ghi_chu": "Ghi chú viết tay trong gia phả: Mất ngày 24/05 Bính Thân (2016) tại Đức Trọng, Lâm Đồng. Sinh: Long, Phương, Kiều, Thu, Nga."
    },
    {
        "stt": 58, "ma": "C2-D6-10", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Bốn", "huy": "Tên gọi: Uyển", "gioi_tinh": "Nam",
        "nam_sinh": "1930 (Canh Ngọ)", "nam_mat": "1993 (Quý Dậu) - Thọ 64",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Con)", "me": "Phạm Thị Thiện",
        "phoi_ngau": "Phan Thị Thao (sinh 1932)",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Huy, Bạch Hưng Uyển, Bạch Hưng Triển, Bạch Thị Hoa, Bạch Thị Lan, Bạch Thị Thủy."
    },
    {
        "stt": 59, "ma": "C2-D6-11", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Lực", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1931 (Tân Mùi)", "nam_mat": "2009 (Kỷ Sửu) - Thọ 79",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Con)", "me": "Phạm Thị Thiện",
        "phoi_ngau": "Ngô Thị Thái (1934 - mất 06/01/2016 Bính Thân, tức 13/02/2016 DL)",
        "mo_phan": "",
        "ghi_chu": "Sinh 8 con trai: Lịch, Thịnh, Vĩnh, Viễn, Mậu, Phượng, Trường, Lường."
    },
    {
        "stt": 60, "ma": "C2-D6-12", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Xuân", "huy": "Liệt sỹ", "gioi_tinh": "Nam",
        "nam_sinh": "1932 (Nhâm Thân)", "nam_mat": "1966 (Bính Ngọ) - Hy sinh",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Con)", "me": "Phạm Thị Thiện",
        "phoi_ngau": "Phạm Thị Vị (sinh 1932)",
        "mo_phan": "Nghĩa trang Liệt sỹ",
        "ghi_chu": "Liệt sỹ chống Mỹ. Sinh: Bạch Hưng Cầm, Bạch Hưng Quyền, Bạch Thị Thu."
    },
    # Đời 7 & 8 Can 2
    {
        "stt": 61, "ma": "C2-D7-01", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thái", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1952 (Nhâm Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Mậu",
        "phoi_ngau": "Nguyễn Thị Bình (sinh 1953)",
        "mo_phan": "",
        "ghi_chu": "Đại tá Quân đội Nhân dân Việt Nam. Ban biên tập Gia phả năm 2014. Con: Bạch Hưng Thắng (1978, có con: Ngọc Mai, Gia Bảo), Bạch Thị Thái Bình (1980)."
    },
    {
        "stt": 62, "ma": "C2-D7-02", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phú", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1956 (Bính Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Mậu",
        "phoi_ngau": "Hoàng Thị Minh (sinh 1960)",
        "mo_phan": "",
        "ghi_chu": "Ban biên tập Gia phả năm 2014; người phụng tự Tổ Can 4 Cụ Kế. Con: Bạch Hưng Đông (1983, có con: Minh, Tuệ Lâm), Bạch Thị Hồng (1985)."
    },
    {
        "stt": 63, "ma": "C2-D7-03", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Dương", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1960 (Canh Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Mậu",
        "phoi_ngau": "Phan Thị Hiền (sinh 1963)",
        "mo_phan": "",
        "ghi_chu": "Con thứ 3 Liệt sỹ Đình. Con: Bạch Hưng Toàn (1988), Bạch Hưng Tùng (1990)."
    },
    {
        "stt": 64, "ma": "C2-D7-04", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tuyên", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Xuân)", "me": "Nguyễn Thị Sen",
        "phoi_ngau": "Đặng Thị Nga",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Quân, Bạch Thị Trang."
    },
    {
        "stt": 65, "ma": "C2-D7-05", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trung", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Xuân)", "me": "Nguyễn Thị Sen",
        "phoi_ngau": "Nguyễn Thị Mai",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Hoàng."
    },
    {
        "stt": 66, "ma": "C2-D7-06", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Vinh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tứ", "me": "Đậu Thị Lộc",
        "phoi_ngau": "Lê Thị Loan",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Dũng."
    },
    {
        "stt": 67, "ma": "C2-D7-07", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Quang", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Niêm", "me": "Lê Thị Mai",
        "phoi_ngau": "Phạm Thị Lệ",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Nam."
    },
    {
        "stt": 68, "ma": "C2-D7-08", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trung", "huy": "Con ông Công", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Công", "me": "Hoàng Thị Thanh",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai ông Công."
    },
    {
        "stt": 69, "ma": "C2-D7-09", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Yêm)", "me": "Phan Thị Lài",
        "phoi_ngau": "Trần Thị Tuyết",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Việt."
    },
    {
        "stt": 70, "ma": "C2-D7-10", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Yêm)", "me": "Phan Thị Lài",
        "phoi_ngau": "Lê Thị Hạnh",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Cường."
    },
    {
        "stt": 71, "ma": "C2-D7-11", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thông", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Vỹ", "me": "Trần Thị Tuyết",
        "phoi_ngau": "Võ Thị Hương",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Huy."
    },
    {
        "stt": 72, "ma": "C2-D7-12", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tiến", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Vỹ", "me": "Trần Thị Tuyết",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con thứ 2 ông Vỹ."
    },
    {
        "stt": 73, "ma": "C2-D7-13", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Ty", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Dị)", "me": "Phạm Thị Lan",
        "phoi_ngau": "Đỗ Thị Thơm",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Long."
    },
    {
        "stt": 74, "ma": "C2-D7-14", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tâm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy (Dị)", "me": "Phạm Thị Lan",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con thứ 2 ông Nậy (Dị)."
    },
    {
        "stt": 75, "ma": "C2-D7-15", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Long", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1948 (Mậu Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cường", "me": "Hoàng Thị Dung",
        "phoi_ngau": "Lê Thị Oanh (sinh 1950)",
        "mo_phan": "",
        "ghi_chu": "Định cư tại Đức Trọng, Lâm Đồng. Con: Tuấn (1976), Vũ (1980), Cúc (1978)."
    },
    {
        "stt": 76, "ma": "C2-D7-16", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phương", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1954 (Giáp Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cường", "me": "Hoàng Thị Dung",
        "phoi_ngau": "Nguyễn Thị Huệ (sinh 1957)",
        "mo_phan": "",
        "ghi_chu": "Con thứ 2 ông Cường. Con: Hải (1982), Hoàng (1986)."
    },
    {
        "stt": 77, "ma": "C2-D7-17", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Kiều", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cường", "me": "Hoàng Thị Dung",
        "phoi_ngau": "Hoàng Thị Hằng (sinh 1961)",
        "mo_phan": "",
        "ghi_chu": "Con thứ 3 ông Cường. Con: Sơn (1985), Hà (1988)."
    },
    {
        "stt": 78, "ma": "C2-D7-18", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Huy", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1956 (Bính Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Phan Thị Thao",
        "phoi_ngau": "Trần Thị Mai (sinh 1959)",
        "mo_phan": "",
        "ghi_chu": "Con cả ông Bốn. Con: Bạch Hưng Hùng (1984), Bạch Thị Nga (1987)."
    },
    {
        "stt": 79, "ma": "C2-D7-19", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Uyển", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1962 (Nhâm Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Phan Thị Thao",
        "phoi_ngau": "Lê Thị Kim (sinh 1965)",
        "mo_phan": "",
        "ghi_chu": "Ban biên tập Gia phả năm 2014. Con: Bạch Hưng Cường (1990), Bạch Thị Trang (1993)."
    },
    {
        "stt": 80, "ma": "C2-D7-20", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Triển", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1968 (Mậu Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Phan Thị Thao",
        "phoi_ngau": "Nguyễn Thị Thanh (sinh 1972)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Kiên (1996), Bạch Hưng Toàn (2000)."
    },
    {
        "stt": 81, "ma": "C2-D7-21", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lịch", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1955 (Ất Mùi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Đinh Thị Mai (sinh 1958)",
        "mo_phan": "",
        "ghi_chu": "Con cả ông Lực. Con: Bạch Hưng Hoàng (1982), Bạch Thị Hồng (1985)."
    },
    {
        "stt": 82, "ma": "C2-D7-22", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thịnh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Vũ Thị Hợi (sinh 1961)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Đạt (1986), Bạch Hưng Thành (1990)."
    },
    {
        "stt": 83, "ma": "C2-D7-23", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Vĩnh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1961 (Tân Sửu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Nguyễn Thị Dung (sinh 1964)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Tuấn (1989), Bạch Hưng Tú (1993)."
    },
    {
        "stt": 84, "ma": "C2-D7-24", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Viễn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Phan Thị Hòa (sinh 1967)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Nam (1992), Bạch Thị Hằng (1995)."
    },
    {
        "stt": 85, "ma": "C2-D7-25", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Mậu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1966 (Bính Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Đậu Thị Lý (sinh 1970)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Sơn (1995), Bạch Hưng Thắng (1999)."
    },
    {
        "stt": 86, "ma": "C2-D7-26", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phượng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1968 (Mậu Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Trần Thị Lan (sinh 1972)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Long (1997), Bạch Thị Linh (2001)."
    },
    {
        "stt": 87, "ma": "C2-D7-27", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trường", "huy": "Con ông Lực", "gioi_tinh": "Nam",
        "nam_sinh": "1971 (Tân Hợi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Lê Thị Tuyết (sinh 1975)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Lâm (2000)."
    },
    {
        "stt": 88, "ma": "C2-D7-28", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1974 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Ngô Thị Thái",
        "phoi_ngau": "Nguyễn Thị Hằng (sinh 1978)",
        "mo_phan": "",
        "ghi_chu": "Con: Bạch Hưng Kiệt (2003)."
    },
    {
        "stt": 89, "ma": "C2-D7-29", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Cầm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xuân", "me": "Phạm Thị Vị",
        "phoi_ngau": "Nguyễn Thị Nga (sinh 1962)",
        "mo_phan": "",
        "ghi_chu": "Con cả Liệt sỹ Xuân. Con: Bạch Thị Mai Thảo (1985), Bạch Hoài Anh (1987), Bạch Nguyễn Trà My (1994), Bạch Nguyễn Thuý Hiền (1997), Bạch Bình Minh (2004 - HC Đồng Olympic Tiếng Anh Quốc gia 2015)."
    },
    {
        "stt": 90, "ma": "C2-D7-30", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Quyền", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xuân", "me": "Phạm Thị Vị",
        "phoi_ngau": "Nguyễn Thị Tuyết (sinh 1969, quê Hưng Đạo, XKLĐ Anh)",
        "mo_phan": "",
        "ghi_chu": "Bảo vệ Làng SOS Vinh, cư trú P. Hà Huy Tập, TP Vinh. Con: Bạch Thị Liên (1993, SV HV Bưu chính), Bạch Hưng Khang (1996)."
    },

    # =========================================================================
    # CAN 3: TỔ CỤ BẠCH HƯNG NGHI
    # =========================================================================
    {
        "stt": 91, "ma": "C3-D3-01", "can": "Can 3", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Nghi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1829 (Kỷ Sửu)", "nam_mat": "Thọ 89 tuổi",
        "ngay_gio": "10/04 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Hồ Thị Phong (thọ 81, giỗ 14/11 ÂL); Bà 2: Nguyễn Thị Ngụ (húy Thông, thọ 36, giỗ 30/06 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 3. Thủ chỉ xã Phú Điền, uy tín lớn, đức độ. Sinh 1 gái (Bạch Thị Tâm) và 1 trai (Bạch Hưng Truyện)."
    },
    {
        "stt": 92, "ma": "C3-D4-01", "can": "Can 3", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Truyện", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất năm 28 tuổi",
        "ngay_gio": "11/07 ÂL", "cha": "Bạch Hưng Nghi", "me": "Nguyễn Thị Ngụ",
        "phoi_ngau": "Nguyễn Thị Năm (người Mệ Giụ, sau cải giá)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Thông minh, học thức cao, 23 tuổi làm Lý trưởng; 28 tuổi khẳng khái chống lệnh bất chính nên bị Cai tổng hãm hại. Sinh 2 con trai: Bạch Hưng Nguyễn, Bạch Hưng Khuyến."
    },
    {
        "stt": 93, "ma": "C3-D5-01", "can": "Can 3", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Nguyễn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "30/12 ÂL", "cha": "Bạch Hưng Truyện", "me": "Nguyễn Thị Năm",
        "phoi_ngau": "Ngô Thị Đôi (thọ 66, giỗ 15/02 ÂL, mộ núi Khe xứ)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai trưởng Cụ Truyện. Sinh 4 trai (Hai, Tư, Niêm, Sáu) và 3 gái (Chắt, Tam, Nhiu)."
    },
    {
        "stt": 94, "ma": "C3-D5-02", "can": "Can 3", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Khuyến", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 66 tuổi",
        "ngay_gio": "06/09 ÂL", "cha": "Bạch Hưng Truyện", "me": "Nguyễn Thị Năm",
        "phoi_ngau": "Hồ Thị Em (thọ 73, giỗ 26/07 ÂL, mộ núi Khe xứ)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Cai xã Phú Điền. Sinh 2 trai (Hiến, Hưng) và 4 gái (Đị, Em, Bưởi, Cọt)."
    },
    # Đời 6 Can 3
    {
        "stt": 95, "ma": "C3-D6-01", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hai", "huy": "Húy Nhưng (Kiểm Nhưng)", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 54 tuổi",
        "ngay_gio": "20/08 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Phạm Thị Đị (giỗ 15/04 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Giữ chức Hương Kiểm của làng. Sinh: Bạch Hưng Thành, Bạch Thị Hoàn, Bạch Thị Sâm."
    },
    {
        "stt": 96, "ma": "C3-D6-02", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Tư", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1906 (Bính Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Phạm Thị Hai (sinh 1914 Giáp Dần, Phú Vệ 1)",
        "mo_phan": "",
        "ghi_chu": "Sinh: Bạch Hưng Liên (1943), Bạch Thị Chắt, Bạch Thị Hồng, Bạch Thị Mạo."
    },
    {
        "stt": 97, "ma": "C3-D6-03", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Niêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi chưa lấy vợ",
        "ngay_gio": "15/05 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Chưa lập gia thất", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 3 Cụ Nguyễn. Ông Bạch Hưng Liên và các con phụng tự."
    },
    {
        "stt": 98, "ma": "C3-D6-04", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Sáu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất ở tuổi 35",
        "ngay_gio": "28/06 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Vợ cải giá", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 4 Cụ Nguyễn. Chưa có con thì mất."
    },
    {
        "stt": 99, "ma": "C3-D6-05", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hiến", "huy": "Đầu phiên Hiến", "gioi_tinh": "Nam",
        "nam_sinh": "1906 (Bính Ngọ)", "nam_mat": "1986 (Bính Dần) - Thọ 82",
        "ngay_gio": "15/02 ÂL", "cha": "Bạch Hưng Khuyến", "me": "Hồ Thị Em",
        "phoi_ngau": "Cao Thị Hoét (thọ 40, giỗ 01/08 ÂL, mộ núi Khe xứ)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Giữ chức Đầu phiên của làng, đức độ, uy tín. Sinh: Bạch Hưng Văn, Bạch Hưng Minh, Bạch Thị Lan."
    },
    {
        "stt": 100, "ma": "C3-D6-06", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hưng", "huy": "Húy Kiến", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất lúc 14 tuổi",
        "ngay_gio": "12/03 ÂL", "cha": "Bạch Hưng Khuyến", "me": "Hồ Thị Em",
        "phoi_ngau": "Chưa lập gia thất", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ Cụ Khuyến. Con ông Hiến thờ tự."
    },
    # Đời 7 & 8 Can 3
    {
        "stt": 101, "ma": "C3-D7-01", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thành", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1936 (Bính Tý)", "nam_mat": "Mất ở tuổi 49",
        "ngay_gio": "12/07 ÂL", "cha": "Bạch Hưng Hai", "me": "Phạm Thị Đị",
        "phoi_ngau": "Hồ Thị Tý (sinh 1936 Bính Tý)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Bí thư Đoàn xã, Cán bộ MTTQ huyện Hưng Nguyên, Cán bộ Ban GT 470, Huân chương Kháng chiến Hạng Nhất. Sinh 5 trai (Chiến, Sỹ, Hương, Thủy, Triều) và 4 gái (Hòa, Nữ, Trinh, Tâm)."
    },
    {
        "stt": 102, "ma": "C3-D7-02", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Liên", "huy": "Thương binh 1/4", "gioi_tinh": "Nam",
        "nam_sinh": "1943 (Quý Mùi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tư", "me": "Phạm Thị Hai",
        "phoi_ngau": "Hồ Thị Miêng (sinh 1941 Tân Tỵ)",
        "mo_phan": "",
        "ghi_chu": "Nhập ngũ 1964, chiến đấu Đường 9 Nam Lào; lái xe vận tải, bị thương hỏng cả 2 mắt. Thương binh nặng 1/4. Huân chương Kháng chiến Hạng 2, Huân chương Chiến sĩ vẻ vang Hạng 2. Sinh: Thắng (1970, có con Phú Vinh, Hồng Quang), Đức (1973 - mất trẻ), Bạch Thị Lài (1979)."
    },
    {
        "stt": 103, "ma": "C3-D7-03", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Văn", "huy": "Thương binh 2/4", "gioi_tinh": "Nam",
        "nam_sinh": "1945 (Ất Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hiến", "me": "Cao Thị Hoét",
        "phoi_ngau": "Trần Thị Kiều (sinh 1948 Mậu Tý, quê Thái Bình)",
        "mo_phan": "",
        "ghi_chu": "Bộ đội pháo cao xạ, chiến đấu Bắc Lào, Đường 9, Đường 20; thương binh mất bàn tay trái. Huân chương Kháng chiến Hạng 2, HC Chiến sĩ vẻ vang Hạng 2. Ban biên tập Gia phả năm 2014. Sinh: Giang (1971), Tình (1979 - ĐH Sư phạm Toán), Yến (1973), Oanh (1975), Nhung (1978)."
    },
    {
        "stt": 104, "ma": "C3-D7-04", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Minh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1948 (Mậu Tý)", "nam_mat": "1993 (Quý Dậu) - Thọ 47",
        "ngay_gio": "04/03 ÂL", "cha": "Bạch Hưng Hiến", "me": "Cao Thị Hoét",
        "phoi_ngau": "Võ Thị Xuân (người làng Văn Viên, Hưng Khánh)",
        "mo_phan": "",
        "ghi_chu": "Bộ đội Trung đoàn 270 tại Quảng Trị, Kế toán trưởng Cty Vật tư Nông nghiệp Nghệ An, cử nhân Kinh tế, Đảng viên 1967. HC Chiến công, HC Kháng chiến Hạng 3. Sinh: Sơn (1976), Hải (1977 - Thượng úy QK4), Bạch Thị Hiền (1987)."
    },

    # =========================================================================
    # CAN 4: TỔ CỤ BẠCH HƯNG KẾ
    # =========================================================================
    {
        "stt": 105, "ma": "C4-D3-01", "can": "Can 4", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Kế", "huy": "Húy Cơ", "gioi_tinh": "Nam",
        "nam_sinh": "1854 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Ngô Thị Tam (húy Chư)",
        "mo_phan": "Xứ Thồ Lô",
        "ghi_chu": "Khởi tổ Can 4. Con trai thứ 4 Cụ Lưu. Con trai Bạch Hưng Thế mất năm 13 tuổi, 2 con gái Mây và Út. Vô tự, được chi họ phụng dưỡng. Trước đây ông Bạch Hưng Đình thờ tự, nay con thứ 2 là ông Bạch Hưng Phú tiếp nối thờ tự."
    }
]

# -----------------------------------------------------------------------------
# 1. TẠO FILE EXCEL: Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx
# -----------------------------------------------------------------------------
print("Đang tạo file Excel: Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx...")
wb = openpyxl.Workbook()

# Style definitions
font_title = Font(name="Times New Roman", size=15, bold=True, color="1F497D")
font_subtitle = Font(name="Times New Roman", size=11, italic=True, color="595959")
font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Times New Roman", size=11, color="000000")
font_bold = Font(name="Times New Roman", size=11, bold=True, color="000000")

fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
fill_subtotal = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF')
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

columns_def = [
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
    ("Cha", 20, align_left),
    ("Mẹ", 20, align_left),
    ("Phối ngẫu (Vợ / Chồng)", 32, align_left),
    ("Mộ phần / An táng", 20, align_left),
    ("Công trạng / Chức vụ / Ghi chú", 45, align_left),
]

def format_sheet_headers(ws, title_text, col_defs):
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(col_defs))
    c1 = ws.cell(row=1, column=1, value=title_text)
    c1.font = font_title
    c1.alignment = align_center

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(col_defs))
    c2 = ws.cell(row=2, column=1, value="TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2 - BẠCH TỘC ĐẠI TÔN (HƯNG PHÚ, HƯNG NGUYÊN)")
    c2.font = font_subtitle
    c2.alignment = align_center

    for col_idx, (col_name, col_width, col_align) in enumerate(col_defs, start=1):
        cell = ws.cell(row=4, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 26

# Sheet 1: Toàn bộ nhân đinh
ws_all = wb.active
ws_all.title = "Toàn bộ nhân đinh"
format_sheet_headers(ws_all, "DANH SÁCH NHÂN ĐINH DÒNG HỌ BẠCH", columns_def)

row_idx = 5
for r in records:
    vals = [
        r["stt"], r["ma"], r["can"], r["doi"], r["ho_ten"], r["huy"],
        r["gioi_tinh"], r["nam_sinh"], r["nam_mat"], r["ngay_gio"],
        r["cha"], r["me"], r["phoi_ngau"], r["mo_phan"], r["ghi_chu"]
    ]
    is_zebra = (row_idx % 2 == 0)
    for c_idx, val in enumerate(vals, start=1):
        c = ws_all.cell(row=row_idx, column=c_idx, value=val)
        c.font = font_bold if c_idx == 5 else font_body
        c.border = thin_border
        c.alignment = columns_def[c_idx-1][2]
        if is_zebra:
            c.fill = fill_zebra
    ws_all.row_dimensions[row_idx].height = 22
    row_idx += 1

# Sheet 2: Can 1
ws_c1 = wb.create_sheet(title="Can 1 - Tổ Cư")
format_sheet_headers(ws_c1, "DANH BẠ CAN 1 - TỔ BẠCH HƯNG CƯ", columns_def)
row_idx = 5
sub_stt = 1
for r in records:
    if r["can"] in ["Can 1", "Khởi phát", "Tiểu Chi 1"]:
        vals = [
            sub_stt, r["ma"], r["can"], r["doi"], r["ho_ten"], r["huy"],
            r["gioi_tinh"], r["nam_sinh"], r["nam_mat"], r["ngay_gio"],
            r["cha"], r["me"], r["phoi_ngau"], r["mo_phan"], r["ghi_chu"]
        ]
        is_zebra = (row_idx % 2 == 0)
        for c_idx, val in enumerate(vals, start=1):
            c = ws_c1.cell(row=row_idx, column=c_idx, value=val)
            c.font = font_bold if c_idx == 5 else font_body
            c.border = thin_border
            c.alignment = columns_def[c_idx-1][2]
            if is_zebra:
                c.fill = fill_zebra
        ws_c1.row_dimensions[row_idx].height = 22
        row_idx += 1
        sub_stt += 1

# Sheet 3: Can 2
ws_c2 = wb.create_sheet(title="Can 2 - Tổ Khâu")
format_sheet_headers(ws_c2, "DANH BẠ CAN 2 - TỔ BẠCH HƯNG KHÂU", columns_def)
row_idx = 5
sub_stt = 1
for r in records:
    if r["can"] in ["Can 2", "Khởi phát", "Tiểu Chi 1"]:
        vals = [
            sub_stt, r["ma"], r["can"], r["doi"], r["ho_ten"], r["huy"],
            r["gioi_tinh"], r["nam_sinh"], r["nam_mat"], r["ngay_gio"],
            r["cha"], r["me"], r["phoi_ngau"], r["mo_phan"], r["ghi_chu"]
        ]
        is_zebra = (row_idx % 2 == 0)
        for c_idx, val in enumerate(vals, start=1):
            c = ws_c2.cell(row=row_idx, column=c_idx, value=val)
            c.font = font_bold if c_idx == 5 else font_body
            c.border = thin_border
            c.alignment = columns_def[c_idx-1][2]
            if is_zebra:
                c.fill = fill_zebra
        ws_c2.row_dimensions[row_idx].height = 22
        row_idx += 1
        sub_stt += 1

# Sheet 4: Can 3 & Can 4
ws_c34 = wb.create_sheet(title="Can 3 & Can 4")
format_sheet_headers(ws_c34, "DANH BẠ CAN 3 & CAN 4 - TỔ BẠCH HƯNG NGHI & TỔ BẠCH HƯNG KẾ", columns_def)
row_idx = 5
sub_stt = 1
for r in records:
    if r["can"] in ["Can 3", "Can 4", "Khởi phát", "Tiểu Chi 1"]:
        vals = [
            sub_stt, r["ma"], r["can"], r["doi"], r["ho_ten"], r["huy"],
            r["gioi_tinh"], r["nam_sinh"], r["nam_mat"], r["ngay_gio"],
            r["cha"], r["me"], r["phoi_ngau"], r["mo_phan"], r["ghi_chu"]
        ]
        is_zebra = (row_idx % 2 == 0)
        for c_idx, val in enumerate(vals, start=1):
            c = ws_c34.cell(row=row_idx, column=c_idx, value=val)
            c.font = font_bold if c_idx == 5 else font_body
            c.border = thin_border
            c.alignment = columns_def[c_idx-1][2]
            if is_zebra:
                c.fill = fill_zebra
        ws_c34.row_dimensions[row_idx].height = 22
        row_idx += 1
        sub_stt += 1

# Sheet 5: Bảng Ngày Kỵ Giỗ
ws_gio = wb.create_sheet(title="Bảng Kỵ Giỗ")
gio_cols = [
    ("STT", 6, align_center),
    ("Ngày Âm lịch", 18, align_center),
    ("Người được giỗ", 24, align_left),
    ("Đời / Can", 18, align_center),
    ("Hưởng thọ / Năm mất", 22, align_center),
    ("Mộ phần / Nơi an táng", 22, align_left),
    ("Ghi chú kỵ giỗ", 35, align_left),
]
format_sheet_headers(ws_gio, "BẢNG KỴ GIỖ TỔ TIÊN VÀ TIỀN NHÂN HỌ BẠCH THEO THỨ TỰ ÂM LỊCH", gio_cols)

gio_list = [
    (1, "10 tháng Giêng", "Cụ bà Phạm Thị Biềng", "Đời 2", "Thọ -", "Xứ Thượng Phù", "Vợ Cụ Bạch Hưng Lưu"),
    (2, "18 tháng Giêng", "Cụ Bạch Hưng Phù (Tư Lang)", "Can 2 - Đời 4", "Thọ 51 tuổi (1912)", "Núi Khe xứ", "Con trai thứ 4 Cụ Khâu"),
    (3, "20 tháng Giêng", "Cụ Bạch Hưng Thấu", "Can 2 - Đời 4", "Mất nhỏ", "Xứ Đồng Mồ", "Con trưởng Cụ Khâu"),
    (4, "20 tháng Giêng", "Cụ Bạch Hưng Em (Con)", "Can 2 - Đời 5", "Thọ 82 tuổi (1981)", "Núi Khe xứ", "Con Cụ Phù"),
    (5, "28 tháng Giêng", "Cụ bà Bùi Thị Hàng Hai", "Khởi phát", "Thọ -", "Xứ Khe Dài", "Vợ hai Cụ Bạch Hưng Hóa"),
    (6, "08 tháng Hai", "Cụ bà Nguyễn Thị Máy", "Can 1 - Đời 3", "Thọ -", "Cồn Giữa xứ", "Vợ cả Cụ Bạch Hưng Cư"),
    (7, "15 tháng Hai", "Cụ bà Ngô Thị Đôi", "Can 3 - Đời 5", "Thọ 66 tuổi", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Nguyễn"),
    (8, "15 tháng Hai", "Cụ Bạch Hưng Hiến (Đầu phiên)", "Can 3 - Đời 6", "Thọ 82 tuổi (1986)", "Núi Khe xứ", "Con cả Cụ Khuyến"),
    (9, "29 tháng Hai", "Cụ Bạch Hưng Lưu (húy Túc)", "Đời 2", "Thọ 76 tuổi (1873)", "Xứ Thượng Phù", "Khởi đầu 4 Can"),
    (10, "04 tháng Ba", "Ông Bạch Hưng Minh", "Can 3 - Đời 7", "Thọ 47 tuổi (1993)", "", "Kế toán trưởng Cty Vật tư NN"),
    (11, "12 tháng Ba", "Bạch Hưng Hưng (húy Kiến)", "Can 3 - Đời 6", "Mất lúc 14 tuổi", "Núi Khe xứ", "Con Cụ Khuyến"),
    (12, "19 tháng Ba", "Cụ bà Phạm Thị Thiện", "Can 2 - Đời 5", "Thọ 67 tuổi", "", "Vợ Cụ Bạch Hưng Em (Con)"),
    (13, "28 tháng Ba", "Cụ bà Nguyễn Quý Nương (Vân)", "Can 1 - Đời 3", "Thọ 85 tuổi", "Xứ Đồng Trọt Bàn", "Vợ ba Cụ Bạch Hưng Cư"),
    (14, "10 tháng Tư", "Cụ Bạch Hưng Nghi (Thủ chỉ)", "Can 3 - Đời 3", "Thọ 89 tuổi", "Núi Khe xứ", "Khởi tổ Can 3"),
    (15, "14 tháng Tư", "Cụ Bạch Hưng Hóa", "Khởi phát", "Thọ 73 tuổi (1802)", "Xứ Cồn Xanh", "Tổ Khởi phát Trung Chi 2"),
    (16, "15 tháng Tư", "Cụ bà Phạm Thị Đị", "Can 3 - Đời 6", "Thọ -", "Núi Khe xứ", "Vợ Cụ Kiểm Nhưng"),
    (17, "21 tháng Tư", "Cụ Bạch Hưng Câu", "Can 2 - Đời 4", "Thọ 31 tuổi (1894)", "", "Con thứ 5 Cụ Khâu"),
    (18, "21 tháng Tư", "Cụ Bạch Hưng Tam (Chắt Mãi)", "Can 2 - Đời 5", "Thọ 60 tuổi (1950)", "", "Con Cụ Xu"),
    (19, "15 tháng Năm", "Ông Bạch Hưng Niêm", "Can 3 - Đời 6", "Mất trẻ (chưa vợ)", "Núi Khe xứ", "Con Cụ Nguyễn"),
    (20, "24 tháng Năm", "Ông Bạch Hưng Cường", "Can 2 - Đời 6", "Thọ 97 tuổi (2016)", "Đức Trọng, Lâm Đồng", "Mất năm Bính Thân 2016"),
    (21, "28 tháng Năm", "Cụ Bạch Hưng Khâm", "Đời 2", "Mất năm 20 tuổi", "Xứ Cồn Vải", "Con trưởng Cụ Tạo"),
    (22, "18 tháng Sáu", "Cụ bà Nguyễn Thị Trúc", "Can 2 - Đời 5", "Thọ 74 tuổi", "", "Vợ Cụ Chắt Mãi"),
    (23, "28 tháng Sáu", "Ông Bạch Hưng Sáu", "Can 3 - Đời 6", "Mất năm 35 tuổi", "Núi Khe xứ", "Con Cụ Nguyễn"),
    (24, "30 tháng Sáu", "Cụ bà Nguyễn Thị Ngụ (Thông)", "Can 3 - Đời 3", "Thọ 36 tuổi", "Núi Khe xứ", "Vợ hai Cụ Bạch Hưng Nghi"),
    (25, "11 tháng Bảy", "Cụ Bạch Hưng Truyện (Lý trưởng)", "Can 3 - Đời 4", "Mất năm 28 tuổi", "Núi Khe xứ", "Con Cụ Nghi, bị hại"),
    (26, "12 tháng Bảy", "Ông Bạch Hưng Thành", "Can 3 - Đời 7", "Thọ 49 tuổi", "Núi Khe xứ", "Huân chương Kháng chiến Hạng 1"),
    (27, "26 tháng Bảy", "Cụ bà Hồ Thị Em", "Can 3 - Đời 5", "Thọ 73 tuổi", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Khuyến"),
    (28, "01 tháng Tám", "Cụ bà Cao Thị Hoét", "Can 3 - Đời 6", "Thọ 40 tuổi", "Núi Khe xứ", "Vợ Cụ Đầu phiên Hiến"),
    (29, "02 tháng Tám", "Cụ Bạch Hưng Khâu (húy Hoài)", "Can 2 - Đời 3", "Thọ 77 tuổi (1897)", "Núi Khe xứ", "Khởi tổ Can 2"),
    (30, "07 tháng Tám", "Cụ bà Bùi Thị Hàng Nhất", "Khởi phát", "Thọ -", "Xứ Cồn Xanh", "Vợ cả Cụ Bạch Hưng Hóa"),
    (31, "07 tháng Tám", "Cụ bà Bùi Thị Muồng", "Can 2 - Đời 3", "Thọ -", "Xứ Rộc Mâm", "Vợ cả Cụ Bạch Hưng Khâu"),
    (32, "11 tháng Tám", "Cụ bà Nguyễn Thị Hàng Hai", "Khởi phát", "Thọ -", "Xứ Khe Dài", "Vợ ba Cụ Bạch Hưng Hóa"),
    (33, "20 tháng Tám", "Cụ Bạch Hưng Hai (Kiểm Nhưng)", "Can 3 - Đời 6", "Thọ 54 tuổi", "Núi Khe xứ", "Hương Kiểm của làng"),
    (34, "28 tháng Tám", "Cụ bà Phạm Thị Hoạt", "Can 2 - Đời 3", "Thọ 62 tuổi", "Xứ Cồn Xanh", "Vợ hai Cụ Bạch Hưng Khâu"),
    (35, "06 tháng Chín", "Cụ Bạch Hưng Khuyến (Cai xã)", "Can 3 - Đời 5", "Thọ 66 tuổi", "Núi Khe xứ", "Cai xã Phú Điền"),
    (36, "28 tháng Chín", "Cụ Bạch Trần Tạo", "Đời 1", "Thọ 30 tuổi (1801)", "Xứ Tuần Tượng", "Khởi tổ Tiểu Chi 1"),
    (37, "08 tháng Mười", "Cụ bà Phạm Thị Hàng Nhất", "Đời 1", "Thọ 76 tuổi", "Xứ Cồn Xanh", "Vợ Cụ Bạch Trần Tạo"),
    (38, "09 tháng Mười", "Cụ Bạch Hưng Cư (húy Khuơ)", "Can 1 - Đời 3", "Thọ -", "Núi Khe xứ", "Khởi tổ Can 1"),
    (39, "14 tháng Mười", "Cụ bà Đậu Thị Diệu", "Can 2 - Đời 4", "Thọ 66 tuổi", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Phù"),
    (40, "22 tháng Mười", "Cụ bà Ngô Thị Hạng Nhị", "Can 1 - Đời 3", "Thọ -", "Cồn Giữa xứ", "Vợ hai Cụ Cư"),
    (41, "22 tháng Mười", "Cụ bà Phan Thị Tác", "Can 2 - Đời 4", "Thọ -", "Núi Khe xứ", "Vợ Cụ Bạch Hưng Xu"),
    (42, "01 tháng Mười Một", "Cụ Bạch Hưng Thưu", "Can 2 - Đời 4", "Vô tự", "Núi Khe xứ", "Con thứ 2 Cụ Khâu"),
    (43, "14 tháng Mười Một", "Cụ bà Hồ Thị Phong", "Can 3 - Đời 3", "Thọ 81 tuổi", "Núi Khe xứ", "Vợ cả Cụ Bạch Hưng Nghi"),
    (44, "06 tháng Chạp (12)", "Cụ bà Ngô Thị Thái", "Can 2 - Đời 6", "Mất 06/01/2016 Bính Thân", "", "Vợ Cụ Bạch Hưng Lực"),
    (45, "30 tháng Chạp (12)", "Cụ Bạch Hưng Nguyễn", "Can 3 - Đời 5", "Thọ -", "Núi Khe xứ", "Con trưởng Cụ Truyện")
]

row_idx = 5
for item in gio_list:
    is_zebra = (row_idx % 2 == 0)
    for c_idx, val in enumerate(item, start=1):
        c = ws_gio.cell(row=row_idx, column=c_idx, value=val)
        c.font = font_bold if c_idx in [2, 3] else font_body
        c.border = thin_border
        c.alignment = gio_cols[c_idx-1][2]
        if is_zebra:
            c.fill = fill_zebra
    ws_gio.row_dimensions[row_idx].height = 22
    row_idx += 1

wb.save(EXCEL_PATH)
print("Đã tạo xong file Excel thành công!")


# -----------------------------------------------------------------------------
# 2. TẠO FILE WORD: Gia_Pha_Ho_Bach.docx
# -----------------------------------------------------------------------------
print("Đang tạo file Word: Gia_Pha_Ho_Bach.docx...")
doc = Document()

# Set page margins
sections = doc.sections
for section in sections:
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

# Colors
COLOR_PRIMARY = RGBColor(31, 73, 125)   # Navy blue
COLOR_SECONDARY = RGBColor(89, 89, 89) # Dark grey
COLOR_TEXT = RGBColor(38, 38, 38)      # Charcoal

# Helper formatting functions
def add_title(text, subtitle=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    
    if subtitle:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run(subtitle)
        run2.font.name = "Times New Roman"
        run2.font.size = Pt(13)
        run2.font.bold = True
        run2.font.color.rgb = COLOR_SECONDARY

def add_h1(text):
    h = doc.add_heading(level=1)
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY

def add_h2(text):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY

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
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
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

# --- TRANG BÌA ---
for _ in range(3):
    doc.add_paragraph()

p_cov1 = doc.add_paragraph()
p_cov1.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_cov1.add_run("BẠCH TỘC ĐẠI TÔN")
r.font.name = "Times New Roman"
r.font.size = Pt(20)
r.font.bold = True
r.font.color.rgb = COLOR_SECONDARY

p_cov2 = doc.add_paragraph()
p_cov2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_cov2.add_run("TÔN ĐỒ TỰ PHẢ\nTIỂU CHI 1 - TRUNG CHI 2")
r.font.name = "Times New Roman"
r.font.size = Pt(28)
r.font.bold = True
r.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph()

p_cov3 = doc.add_paragraph()
p_cov3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_cov3.add_run("LÀNG PHÚ ĐIỀN - XÃ HƯNG PHÚ\nHUYỆN HƯNG NGUYÊN - TỈNH NGHỆ AN")
r.font.name = "Times New Roman"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_TEXT

for _ in range(6):
    doc.add_paragraph()

p_cov4 = doc.add_paragraph()
p_cov4.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_cov4.add_run("LƯU TRUYỀN HẬU THẾ - NĂM GIÁP NGỌ (2014)")
r.font.name = "Times New Roman"
r.font.size = Pt(12)
r.font.italic = True
r.font.color.rgb = COLOR_SECONDARY

doc.add_page_break()

# --- LỜI NÓI ĐẦU ---
add_h1("LỜI NÓI ĐẦU")
add_callout(
    "\"Con người ta sinh ra ở đời ai cũng có Tổ, có Tông, như cây có cội, như sông có nguồn. "
    "Đó là đạo lý truyền thống ngàn đời tốt đẹp của dân tộc ta.\"",
    "Cội Nguồn Sinh Dưỡng"
)
add_p("Cuốn phả hệ này là Tôn Đồ Tự Phả của Tiểu Chi 1, Trung Chi 2 thuộc Bạch Tộc Đại Tôn, nguyên quán tại làng Phú Điền, xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An.")
add_p("Vào những năm 1990, ông Bạch Hưng Thủng đã có công đầu dày công sưu tầm, khảo cứu từ các văn bản chữ Hán cổ của dòng tộc, dịch thuật sang chữ Quốc ngữ để lưu truyền cho con cháu.")
add_p("Đến năm 2014 (Giáp Ngọ), nhận thấy nhu cầu cập nhật các thế hệ con cháu nối tiếp, đồng thời để bảo tồn nguyên vẹn các di huấn, mộ phần và ngày kỵ giỗ của các bậc tiền nhân, ban tộc biểu gồm các ông: Bạch Hưng Thái, Bạch Hưng Phú, Bạch Hưng Văn, Bạch Hưng Uyển, Bạch Hưng Tùng đã cùng nhau đối chiếu, bổ sung, chỉnh lý và in ấn thành tập sách gia phả hoàn chỉnh này nhằm giúp con cháu:")
add_p("1. Biết rõ cội nguồn, gốc tích và mối quan hệ huyết thống thiêng liêng giữa các nhánh, các Can.")
add_p("2. Nắm vững ngày giỗ Âm lịch, nơi an táng (mộ phần) của tổ tiên để đời đời phụng tự, khói hương chu đáo.")
add_p("3. Noi gương truyền thống hiếu học, yêu nước, kiên trung của các bậc tiền bối, nỗ lực phấn đấu làm rạng danh dòng họ Bạch.")

# --- LỊCH SỬ HỌ BẠCH ---
add_h1("VÀI NÉT VỀ LỊCH SỬ HỌ BẠCH & DANH NHÂN BẠCH LIÊU")
add_h2("1. Thủy tổ Trạng nguyên Bạch Liêu (1236 - 1315)")
add_p("Dòng họ Bạch Việt Nam tự hào có cụ Thủy tổ là Trạng nguyên Bạch Liêu, sinh ngày 15 tháng Giêng năm Bính Thìn (1236) tại làng Yên Xá, huyện Chân Phúc (nay thuộc xã Mã Thành, huyện Yên Thành, tỉnh Nghệ An).")
add_p("Năm Bính Dần (1266), niên hiệu Thiệu Long thứ 9 đời vua Trần Thánh Tông, ngài đỗ Đệ nhất giáp tiến sĩ đệ nhất danh (Trạng nguyên) khi tròn 31 tuổi. Ngài nổi tiếng là bậc thông minh xuất chúng, học rộng tài cao, nhớ lâu và có tài năng kiệt xuất về chính trị, quân sự và bang giao.")
add_p("Sau khi đỗ Trạng nguyên, Thượng phụ Chiêu Minh Đại vương Trần Quang Khải rất mến phục đức tài của ngài, đã mời ngài về làm Môn khách trong phủ. Trong cuộc kháng chiến chống quân xâm lược Nguyên - Mông lần thứ 2 (1285) và lần thứ 3 (1287 - 1288), ngài có công lao to lớn giúp Trần Quang Khải hoạch định cơ mưu chiến lược, thảo hịch đánh giặc, góp phần vào những chiến thắng hiển hách như Chương Dương, Hàm Tử.")
add_p("Ngài cũng là tác giả của \"Biến pháp tam chương\" (Ba chương phép biến đổi), dâng lên vua Trần nhằm cải cách hành chính, quản lý đinh số, ruộng đất và củng cố quân đội, đặt nền móng vững chắc cho nền thịnh trị của nhà Trần. Cụ tạ thế ngày 24 tháng Giêng năm Ất Mão (1315), thọ 80 tuổi. Triều đình truy phong tước hiệu cao quý, nhân dân nhiều nơi lập đền thờ phụng để ghi nhớ công đức.")

add_h2("2. Đền thờ họ Bạch tại Hưng Phú, Hưng Nguyên")
add_p("Tại xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An, con cháu họ Bạch đã lập đền thờ Cụ Trạng nguyên Bạch Liêu và các bậc Tiên tổ. Đền thờ dòng họ Bạch tại Hưng Phú đã được Nhà nước xếp hạng là Di tích Lịch sử - Văn hóa cấp Tỉnh theo Quyết định số 2015/QĐ-UBND ngày 16/12/1993.")
add_p("Đến năm 2013, đền thờ đã được toàn thể con cháu dòng họ chung sức tôn tạo, nâng cấp khang trang, bề thế, trở thành nơi quy tụ tâm linh thiêng liêng của bà con cô bác họ Bạch khắp mọi miền đất nước.")

# --- KHỞI PHÁT & ĐỜI 1, ĐỜI 2 ---
add_h1("TỔ KHỞI PHÁT TRUNG CHI 2 & ĐỜI THỨ NHẤT, THỨ HAI")
add_h2("Tổ Khởi Phát Trung Chi 2: Cụ BẠCH HƯNG HÓA")
add_p("Canh Tuất (1730) - Mất: Nhâm Tuất (1802), hưởng thọ 73 tuổi.", "Sinh: ")
add_p("Ngày 14 tháng 4 Âm lịch.", "Ngày kỵ giỗ: ")
add_p("Xứ Cồn Xanh.", "Nơi an táng: ")
add_p("Bà 1: Bùi Thị Hàng Nhất (giỗ 07/08 ÂL, mộ Cồn Xanh); Bà 2: Bùi Thị Hàng Hai (giỗ 28/01 ÂL, mộ Khe Dài); Bà 3: Nguyễn Thị Hàng Hai (giỗ 11/08 ÂL, mộ Khe Dài).", "Phối ngẫu: ")
add_p("Sinh được 8 người con: Bạch Trần Tạo (lập Tiểu Chi 1), Bạch Trần Điển (lập Tiểu Chi 2), Bạch Trần Khơm (lập Tiểu Chi 3), Bạch Khắc Chẩn, Bạch Hưng Đại, Bạch Thị Thao, Bạch Thị Đỏ...", "Hậu duệ: ")

add_h2("Đời Thứ Nhất (Tiểu Chi 1): Cụ BẠCH TRẦN TẠO")
add_p("Nhâm Thìn (1772) - Mất: Tân Dậu (1801), hưởng thọ 30 tuổi.", "Sinh: ")
add_p("Ngày 28 tháng 9 Âm lịch.", "Ngày kỵ giỗ: ")
add_p("Xứ Tuần Tượng.", "Nơi an táng: ")
add_p("Cụ bà Phạm Thị Hàng Nhất, thọ 76 tuổi, ngày giỗ 08/10 ÂL, mộ táng tại xứ Cồn Xanh.", "Phối ngẫu: ")
add_p("Sinh được 2 người con trai: Bạch Hưng Khâm, Bạch Hưng Lưu.", "Hậu duệ: ")

add_h2("Đời Thứ Hai: Cụ BẠCH HƯNG KHÂM & Cụ BẠCH HƯNG LƯU")
add_p("Con trai trưởng Cụ Tạo. Mất năm 20 tuổi, chưa lập gia đình. Giỗ ngày 28/05 ÂL, mộ đặt tại xứ Cồn Vải. Nhánh Cụ Lưu phụng tự.", "1. Cụ Bạch Hưng Khâm: ")
add_p("Mậu Ngọ (1798) - Mất: Quý Dậu (1873), thọ 76 tuổi. Tên húy là Túc. Giỗ ngày 29/02 ÂL, mộ táng xứ Thượng Phù. Vợ là Cụ bà Phạm Thị Biềng (giỗ 10/01 ÂL, mộ xứ Thượng Phù).", "2. Cụ Bạch Hưng Lưu: ")
add_p("Cụ Bạch Hưng Lưu sinh được 4 người con trai phân thành 4 Can (Can 1: Cụ Cư; Can 2: Cụ Khâu; Can 3: Cụ Nghi; Can 4: Cụ Kế) và 4 người con gái: Bạch Thị Thuật (lấy chồng Yên Cư), Bạch Thị Khoát (lấy chồng Thần Võ), Bạch Thị Khoan (lấy chồng Phú Điền), Bạch Thị Út (lấy chồng Phú Vệ).")

# --- CAN 1 ---
add_h1("CAN THỨ NHẤT: TỔ CAN BẠCH HƯNG CƯ (Đời 3 đến Đời 8)")
add_h2("Đời Thứ 3: Cụ BẠCH HƯNG CƯ (Húy Khuơ)")
add_p("Sinh năm Kỷ Mão (1819). Là con trai trưởng của Cụ Bạch Hưng Lưu. Ngày giỗ: Mồng 9 tháng 10 ÂL. Mộ táng tại núi Khe xứ.")
add_p("Bà 1: Nguyễn Thị Máy (giỗ 08/02 ÂL, mộ Cồn Giữa xứ); Bà 2: Ngô Thị Hạng Nhị (giỗ 22/10 ÂL, mộ Cồn Giữa xứ); Bà 3: Nguyễn Quý Nương (hiệu Từ Thuần, húy Vân, thọ 85, giỗ 28/03 ÂL, mộ Đồng Trọt Bàn).", "Phối ngẫu: ")
add_p("Con gái: Bạch Thị Hoạch, Khương, Ba, Sáu. Con trai: Hạnh (mất nhỏ), Thiết, Ý, Tơn, Ngân, Lùa (vô tự), Tiu (vô tự).", "Hậu duệ: ")

add_h2("Đời Thứ 4 đến Đời Thứ 7 (Can 1)")
add_p("Thiết (húy Khuyến, mộ núi Khe xứ); Ý (húy Giám, mộ núi Khe xứ); Tơn (húy Cửu, mộ núi Khe xứ); Ngân (húy Khoản, mộ núi Khe xứ); Khôn; Hai.", "Đời 4: ")
add_p("Khớt (húy Nhược), Em (Bâu), Chế (Tam), Cận, Khôn, Hai (Hiểm).", "Đời 5: ")
add_p("Chắt, Chính, Mô, Cháu, Nhã, Mận, Mai, Nhỏ.", "Đời 6: ")
add_p("Bạch Hưng Tùng, Bạch Hưng Hùng, Bạch Hưng Dũng, Bạch Hưng Thất, Bạch Quốc Việt, Bạch Hưng Nam, Bạch Hưng Lập, Bạch Hưng Thanh, Bạch Hưng Diệu, Bạch Hưng Lợi, Bạch Hưng Hải, Bạch Thái Nguyên, Bạch Quang Ngọc.", "Đời 7: ")

# --- CAN 2 ---
add_h1("CAN THỨ HAI: TỔ CAN BẠCH HƯNG KHÂU (Đời 3 đến Đời 8)")
add_h2("Đời Thứ 3: Cụ BẠCH HƯNG KHÂU (Húy Hoài)")
add_p("Tân Tỵ (1821) - Mất: Đinh Dậu (1897), hưởng thọ 77 tuổi. Ngày giỗ: 02/08 ÂL. Mộ táng tại núi Khe xứ.", "Sinh: ")
add_p("Bà 1: Bùi Thị Muồng (giỗ 07/08 ÂL, mộ Rộc Mâm); Bà 2: Phạm Thị Hoạt (thọ 62, giỗ 28/08 ÂL, mộ Cồn Xanh).", "Phối ngẫu: ")
add_p("5 con trai: Thấu (mất nhỏ), Thưu (vô tự), Xu (húy Mâu, sinh 1854), Phù (húy Tư Lang, 1862-1912), Câu (1864-1894); và 1 con gái: Bạch Thị Hai.", "Hậu duệ: ")

add_h2("Các Thế Hệ Tiêu Biểu Can 2 (Đời 5, 6, 7)")
add_p("Liệt sỹ chống Mỹ cứu nước, hy sinh năm 1968 (Mậu Thân) tại chiến trường Quảng Trị. Vợ là Cụ bà Nguyễn Thị Mậu (1923 - 2012, thọ 89 tuổi). Sinh các con: Đại tá Bạch Hưng Thái, Bạch Hưng Phú, Bạch Hưng Dương, Bạch Thị Phượng, Bạch Thị Lương.", "Liệt sỹ Bạch Hưng Đình (1925 - 1968): ")
add_p("Sinh năm Canh Thân (1920). Vợ là Hoàng Thị Dung (sinh 1921). Ghi chú viết tay trong sách: Cụ Cường tạ thế ngày 24/05 năm Bính Thân (2016) tại thị trấn Liên Nghĩa, huyện Đức Trọng, tỉnh Lâm Đồng (thọ 97 tuổi). Sinh: Long, Phương, Kiều, Thu, Nga.", "Cụ Bạch Hưng Cường (1920 - 2016): ")
add_p("Sinh năm 1931 (Tân Mùi), mất năm 2009 (Kỷ Sửu), thọ 79 tuổi. Vợ là Cụ bà Ngô Thị Thái (1934 - tạ thế ngày 06/01/2016 Bính Thân, tức 13/02/2016 DL). Sinh 8 con trai: Lịch, Thịnh, Vĩnh, Viễn, Mậu, Phượng, Trường, Lường.", "Cụ Bạch Hưng Lực (1931 - 2009): ")
add_p("Sinh năm 1932 (Nhâm Thân), hy sinh năm 1966 (Bính Ngọ). Vợ là Phạm Thị Vị (sinh 1932). Sinh: Bạch Hưng Cầm, Bạch Hưng Quyền, Bạch Thị Thu.", "Liệt sỹ Bạch Hưng Xuân (1932 - 1966): ")
add_p("Sinh năm 1958, vợ là Nguyễn Thị Nga (sinh 1962). Các con cháu: Bạch Thị Mai Thảo (1985), Bạch Hoài Anh (1987), Bạch Nguyễn Trà My (1994), Bạch Nguyễn Thuý Hiền (1997), Bạch Bình Minh (2004 - Huy chương Đồng Olympic Tiếng Anh cấp Quốc gia năm 2015).", "Bạch Hưng Cầm (Đời 7): ")
add_p("Sinh năm 1964, công tác tại Làng SOS Vinh, cư trú P. Hà Huy Tập, TP Vinh. Vợ là Nguyễn Thị Tuyết (sinh 1969, XKLĐ Anh). Con: Bạch Thị Liên (1993, HV Bưu chính), Bạch Hưng Khang (1996).", "Bạch Hưng Quyền (Đời 7): ")

# --- CAN 3 ---
add_h1("CAN THỨ BA: TỔ CAN BẠCH HƯNG NGHI (Đời 3 đến Đời 8)")
add_h2("Đời Thứ 3: Cụ BẠCH HƯNG NGHI (Thủ Chỉ Xã Phú Điền)")
add_p("Kỷ Sửu (1829) - Hưởng thọ 89 tuổi. Ngày giỗ: 10/04 ÂL. Lăng mộ đặt tại núi Khe xứ.", "Sinh: ")
add_p("Cụ nguyên là Thủ chỉ đứng đầu xã Phú Điền thời bấy giờ, tính tình cương trực, có học thức và uy tín lớn trong dân làng.", "Công đức: ")
add_p("Bà cả: Hồ Thị Phong (thọ 81, giỗ 14/11 ÂL); Bà hai: Nguyễn Thị Ngụ (húy Thông, thọ 36, giỗ 30/06 ÂL, mộ núi Khe xứ).", "Phối ngẫu: ")
add_p("Sinh được 1 gái: Bạch Thị Tâm (lấy Ngô Phú Hậu) và 1 trai: Bạch Hưng Truyện.", "Hậu duệ: ")

add_h2("Đời Thứ 4: Cụ BẠCH HƯNG TRUYỆN (Lý Trưởng Kiên Trung)")
add_p("Con trai Cụ Nghi. Thông minh, học thức cao, năm 23 tuổi được cử làm Lý trưởng. Năm 28 tuổi, vì cương trực không chịu nghe theo lệnh bất chính của quan Cai tổng nên cụ đã bị Cai tổng hãm hại qua đời ngày 11/07 ÂL. Lăng mộ đặt tại núi Khe xứ. Vợ là Cụ bà Nguyễn Thị Năm (người Mệ Giụ, sau cải giá). Sinh 2 con trai: Bạch Hưng Nguyễn, Bạch Hưng Khuyến.")

add_h2("Đời Thứ 5 & Đời Thứ 6 (Can 3)")
add_p("Con trưởng Cụ Truyện. Giỗ 30/12 ÂL, mộ núi Khe xứ. Vợ là Ngô Thị Đôi (thọ 66, giỗ 15/02 ÂL). Sinh: Chắt, Tam, Nhiu, Bạch Hưng Hai, Tư, Niêm, Sáu.", "1. Cụ Bạch Hưng Nguyễn: ")
add_p("Cai xã Phú Điền, thọ 66 tuổi, giỗ 06/09 ÂL. Vợ là Hồ Thị Em (thọ 73, giỗ 26/07 ÂL). Mộ núi Khe xứ. Sinh: Bạch Hưng Hiến, Bạch Hưng Hưng (Kiến, mất 14 tuổi), Bạch Thị Đị, Em, Bưởi, Cọt.", "2. Cụ Bạch Hưng Khuyến: ")
add_p("Cụ Bạch Hưng Hai (Kiểm Nhưng, thọ 54, giỗ 20/08 ÂL); Cụ Bạch Hưng Tư (sinh 1906, vợ Phạm Thị Hai); Cụ Bạch Hưng Niêm (mất trẻ, giỗ 15/05 ÂL); Cụ Bạch Hưng Sáu (mất 35 tuổi, giỗ 28/06 ÂL); Cụ Bạch Hưng Hiến (Đầu phiên, 1906 - 1986, thọ 82, giỗ 15/02 ÂL, vợ Cao Thị Hoét mất 40 tuổi giỗ 01/08 ÂL).", "Các Cụ Đời 6: ")

add_h2("Đời Thứ 7 Tiêu Biểu Can 3")
add_p("Sinh năm 1936 Bính Tý. Bí thư Xã đoàn, Cán bộ MTTQ Huyện, Cán bộ Ban GT 470 Diễn Châu, Huân chương Kháng chiến Hạng Nhất. Giỗ 12/07 ÂL (thọ 49 tuổi). Vợ là Hồ Thị Tý (sinh 1936). Sinh: Chiến, Sỹ, Hương, Thủy, Triều, Hòa, Nữ, Trinh, Tâm.", "Bạch Hưng Thành (1936 - 1985): ")
add_p("Sinh năm 1943 Quý Mùi. Nhập ngũ 1964, chiến đấu Đường 9 Nam Lào; lái xe tiếp vận, bị thương hỏng cả 2 mắt. Thương binh nặng hạng 1/4. Huân chương Kháng chiến Hạng 2. Vợ là Hồ Thị Miêng (sinh 1941). Sinh: Thắng, Đức, Lài.", "Bạch Hưng Liên (Thương binh 1/4): ")
add_p("Sinh năm 1945 Ất Dậu. Bộ đội pháo cao xạ, chiến đấu Bắc Lào, Đường 9, Đường 20; bị thương cụt bàn tay trái. Thương binh hạng 2/4. Huân chương Kháng chiến Hạng 2. Vợ là Trần Thị Kiều (sinh 1948). Sinh: Giang, Tình (GV Toán THPT Quỳ Hợp 3), Yến, Oanh, Nhung.", "Bạch Hưng Văn (Thương binh 2/4): ")
add_p("Sinh năm 1948 Mậu Tý, mất năm 1993 (thọ 47 tuổi, giỗ 04/03 ÂL). Bộ đội Trung đoàn 270 tại Quảng Trị, Kế toán trưởng Cty Vật tư Nông nghiệp Nghệ An, cử nhân Kinh tế, Đảng viên 1967. Vợ là Võ Thị Xuân. Sinh: Sơn, Hải (Thượng úy QK4), Hiền.", "Bạch Hưng Minh (1948 - 1993): ")

# --- CAN 4 ---
add_h1("CAN THỨ TƯ: TỔ CAN BẠCH HƯNG KẾ")
add_p("Giáp Dần (1854). Là con trai thứ 4 của Cụ Bạch Hưng Lưu.", "Sinh năm: ")
add_p("Cụ bà Ngô Thị Tam (tên húy là Chư).", "Phối ngẫu: ")
add_p("Con trai Bạch Hưng Thế mất năm 13 tuổi; 2 con gái là Bạch Thị Mây và Bạch Thị Út.", "Hậu duệ: ")
add_p("Lăng mộ đặt tại xứ Thồ Lô.", "Nơi an táng: ")
add_p("Do không có con trai nối dòng, sinh thời Cụ nhận cháu trong họ phụng dưỡng. Trước đây ông Bạch Hưng Đình đảm nhận việc thờ tự khói hương; nay con thứ hai của ông Đình là ông Bạch Hưng Phú tiếp nối trọng trách thờ tự Cụ tại gia đường.", "Phụng tự: ")

# --- BẢNG TRA CỨU NGÀY GIỖ ---
add_h1("BẢNG TRA CỨU NGÀY GIỖ ÂM LỊCH TOÀN DÒNG HỌ")
table_doc = doc.add_table(rows=1, cols=4)
table_doc.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr_cells = table_doc.rows[0].cells
hdr_titles = ["Ngày Âm lịch", "Người được giỗ", "Đời / Can", "Mộ phần / Nơi an táng"]
for idx, title in enumerate(hdr_titles):
    hdr_cells[idx].text = title
    tcPr = hdr_cells[idx]._tc.get_or_add_tcPr()
    shd = parse_xml(r'<w:shd {} w:fill="1F497D"/>'.format(nsdecls('w')))
    tcPr.append(shd)
    p = hdr_cells[idx].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

for item in gio_list:
    row_cells = table_doc.add_row().cells
    vals = [item[1], item[2], item[3], item[5]]
    for idx, text in enumerate(vals):
        row_cells[idx].text = text
        p = row_cells[idx].paragraphs[0]
        if idx in [0, 2]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            if idx in [0, 1]:
                r.font.bold = True

doc.save(DOCX_PATH)
print("Đã tạo xong file Word thành công!")

