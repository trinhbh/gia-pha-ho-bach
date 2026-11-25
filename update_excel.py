# -*- coding: utf-8 -*-
"""
Script tái thiết lập 100% dữ liệu phả hệ họ Bạch DỰA HOÀN TOÀN VÀO VĂN BẢN (KHÔNG THEO SƠ ĐỒ)
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

records = [
    # KHỞI PHÁT & TIỀN ĐẠI
    {
        "stt": 1, "ma": "KP-01", "can": "Trung Chi 2", "doi": "Khởi phát",
        "ho_ten": "Bạch Hưng Hóa", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1730 (Canh Tuất)", "nam_mat": "1802 (Nhâm Tuất) - Thọ 73",
        "ngay_gio": "14/04 ÂL", "cha": "", "me": "",
        "phoi_ngau": "Bùi Thị Hàng Nhất; Bùi Thị Hàng Hai; Nguyễn Thị Hàng Hai",
        "mo_phan": "Xứ Cồn Xanh",
        "ghi_chu": "Tổ Khởi phát Trung Chi 2 họ Bạch, sinh 8 người con."
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
        "ghi_chu": "Con trưởng Cụ Tạo. Nhánh Cụ Lưu phụng tự."
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
    # CAN 1: TỔ BẠCH HƯNG CƯ (CHUẨN XÁC THEO VĂN BẢN TRANG 20 - 32)
    # =========================================================================
    {
        "stt": 5, "ma": "C1-D3-01", "can": "Can 1", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Cư", "huy": "Húy Khuơ", "gioi_tinh": "Nam",
        "nam_sinh": "1819 (Kỷ Mão)", "nam_mat": "",
        "ngay_gio": "09/10 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Nguyễn Thị Máy (giỗ 08/02 ÂL); Bà 2: Ngô Thị Hạng Nhị; Bà 3: Nguyễn Quý Nương (vợ hai/húy Vân, thọ 85, giỗ 28/03 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 1. Con trai cả Cụ Lưu. Con bà Nhất Ngô Thị Hạng Nhị: Hoạch, Khương, Ba, Sáu, Hạnh, Thiết, Ý. Con bà Nguyễn Thị Vân: Tơn, Ngân, Lùa, Tiu."
    },
    # Đời 4 Can 1
    {
        "stt": 6, "ma": "C1-D4-01", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Hạnh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Chết lúc còn nhỏ",
        "ngay_gio": "", "cha": "Bạch Hưng Cư", "me": "Ngô Thị Hạng Nhị",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai Cụ Cư với bà Ngô Thị Hạng Nhị, mất sớm."
    },
    {
        "stt": 7, "ma": "C1-D4-02", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thiết", "huy": "Húy Khải", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "01/03 ÂL", "cha": "Bạch Hưng Cư", "me": "Ngô Thị Hạng Nhị",
        "phoi_ngau": "Phan Thị Nghi", "mo_phan": "Khe Xứ",
        "ghi_chu": "Con Cụ Cư. Sinh một con trai là Bạch Hưng Khớt (có vợ không con, mất năm 46 tuổi)."
    },
    {
        "stt": 8, "ma": "C1-D4-03", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Ý", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "01/10 ÂL", "cha": "Bạch Hưng Cư", "me": "Ngô Thị Hạng Nhị",
        "phoi_ngau": "", "mo_phan": "Núi Khe xứ (chiêu hồn cải táng)",
        "ghi_chu": "Con thứ 3 bà Ngô Thị Hạng Nhị. Đi lính thời Nguyễn, chết đuối ở sông Hiếu, Nghĩa Đàn."
    },
    {
        "stt": 9, "ma": "C1-D4-04", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Tơn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 60 tuổi",
        "ngay_gio": "12/06 ÂL", "cha": "Bạch Hưng Cư", "me": "Nguyễn Thị Vân",
        "phoi_ngau": "Phạm Thị Phòng (giỗ 07/05 ÂL)", "mo_phan": "",
        "ghi_chu": "Con trai Cụ Cư với bà Vân. Sinh 3 trai 1 gái: Bạch Hưng Em (Bâu), Bạch Hưng Chế (Tam), Bạch Hưng Cận (Tam), Bạch Thị Tiu."
    },
    {
        "stt": 10, "ma": "C1-D4-05", "can": "Can 1", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Ngân", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "30/06 ÂL", "cha": "Bạch Hưng Cư", "me": "Nguyễn Thị Vân",
        "phoi_ngau": "Bà cả: Hồ Thị Niêm (giỗ 15/07 ÂL); Bà hai: Nguyễn Thị Em (húy Tảo, giỗ 17/09 ÂL)",
        "mo_phan": "Rú Sau (cây đa Bò Mẹp)",
        "ghi_chu": "Con Cụ Cư với bà Vân. Bà cả sinh: Bạch Hưng Khôn, Bạch Hưng Hai. Bà hai sinh: Bạch Thị Tứ, Bạch Thị Ngụ."
    },
    # Đời 5 Can 1
    {
        "stt": 11, "ma": "C1-D5-01", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Khớt", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất năm 46 tuổi",
        "ngay_gio": "", "cha": "Bạch Hưng Thiết", "me": "Phan Thị Nghi",
        "phoi_ngau": "Có vợ nhưng không có con", "mo_phan": "",
        "ghi_chu": "Con trai độc nhất của Cụ Bạch Hưng Thiết."
    },
    {
        "stt": 12, "ma": "C1-D5-02", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Húy Bâu", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "15/03 ÂL", "cha": "Bạch Hưng Tơn", "me": "Phạm Thị Phòng",
        "phoi_ngau": "Phạm Thị Em (người Văn Viên)", "mo_phan": "Xứ Khe",
        "ghi_chu": "Con Cụ Tơn. Sinh một con trai là Bạch Hưng Chắt (năm 1945 thất lạc/chết đói)."
    },
    {
        "stt": 13, "ma": "C1-D5-03", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Chế", "huy": "Bạch Hưng Tam", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 74 tuổi",
        "ngay_gio": "19/01 ÂL", "cha": "Bạch Hưng Tơn", "me": "Phạm Thị Phòng",
        "phoi_ngau": "Nguyễn Thị Tứ (1900 - 1987, thọ 88, giỗ 11/10 ÂL, mộ Nghĩa Đức, Nghĩa Đàn)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con Cụ Tơn. Sinh 5 con trai: Bạch Hưng Cháu, Bạch Hưng Nhã, Bạch Hưng Mận, Bạch Hưng Mai, Bạch Hưng Nhỏ; và con gái Bạch Thị Mơ."
    },
    {
        "stt": 14, "ma": "C1-D5-04", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Cận", "huy": "Húy Tam", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tơn", "me": "Phạm Thị Phòng",
        "phoi_ngau": "Nguyễn Thị Mịn (sau cải giá)", "mo_phan": "",
        "ghi_chu": "Con thứ 3 Cụ Tơn. Sinh 1 con gái là Bạch Thị Cháu (lấy chồng chợ Vực)."
    },
    {
        "stt": 15, "ma": "C1-D5-05", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Khôn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Ngân", "me": "Hồ Thị Niêm",
        "phoi_ngau": "Dư Thị Bốn (người Hiệu Mỹ)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con Cụ Ngân. Sinh 1 con trai là Bạch Hưng Chính (mất nhỏ) và 3 con gái (Cháu, Nhỏ, Hòa)."
    },
    {
        "stt": 16, "ma": "C1-D5-06", "can": "Can 1", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Hai", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "23/03 ÂL", "cha": "Bạch Hưng Ngân", "me": "Hồ Thị Niêm",
        "phoi_ngau": "Nguyễn Thị Tứ (giỗ 02/02 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con Cụ Ngân. Sinh 1 con trai là Bạch Hưng Mô (mất nhỏ) và 2 con gái (Cháu, Hữu)."
    },
    # Đời 6 Can 1
    {
        "stt": 17, "ma": "C1-D6-01", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Chắt", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất tích năm 1945",
        "ngay_gio": "", "cha": "Bạch Hưng Em (Bâu)", "me": "Phạm Thị Em",
        "phoi_ngau": "Chưa có vợ con", "mo_phan": "",
        "ghi_chu": "Con trai ông Bạch Hưng Em (Bâu)."
    },
    {
        "stt": 18, "ma": "C1-D6-02", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Chính", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi còn nhỏ",
        "ngay_gio": "", "cha": "Bạch Hưng Khôn", "me": "Dư Thị Bốn",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai ông Bạch Hưng Khôn."
    },
    {
        "stt": 19, "ma": "C1-D6-03", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mô", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi còn nhỏ",
        "ngay_gio": "", "cha": "Bạch Hưng Hai", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai ông Bạch Hưng Hai."
    },
    {
        "stt": 20, "ma": "C1-D6-04", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Cháu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1923 (Quý Hợi)", "nam_mat": "2014 (Giáp Ngọ)",
        "ngay_gio": "30/06 ÂL", "cha": "Bạch Hưng Chế (Tam)", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "Nguyễn Thị Giá (sinh 1923, Yên Cư)", "mo_phan": "Núi Khe Xứ",
        "ghi_chu": "Con trai cả ông Bạch Hưng Chế. Sinh 4 trai: Tùng, Dũng, Hùng, Thất; 3 gái: Thành, Phùng, Tâm."
    },
    {
        "stt": 21, "ma": "C1-D6-05", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nhạ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1923 (Quý Hợi)", "nam_mat": "1998 - Thọ 76",
        "ngay_gio": "23/03 ÂL", "cha": "Bạch Hưng Chế (Tam)", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "Nguyễn Thị Chắt (1926 - 2005, thọ 80)", "mo_phan": "Nghĩa Đức, Nghĩa Đàn",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Chế. Khai hoang Nghĩa Đức. Sinh 2 trai: Việt, Nam; 7 gái: Hòa, Nguyệt, Bảo, Toàn, Thảo, Hiền, Lương."
    },
    {
        "stt": 22, "ma": "C1-D6-06", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mận", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1935 (Ất Hợi)", "nam_mat": "2002 - Thọ 68",
        "ngay_gio": "10/02 ÂL", "cha": "Bạch Hưng Chế (Tam)", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "Nguyễn Thị Loan (1935 - 1972, thọ 38)", "mo_phan": "Nghĩa Đức, Nghĩa Đàn",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Chế. Sinh 4 trai: Lập, Thanh, Diệu, Lợi; 2 gái: Minh, Oanh."
    },
    {
        "stt": 23, "ma": "C1-D6-07", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Mai", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1939 (Kỷ Mão)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chế (Tam)", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "Hồ Thị Tình (sinh 1944)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 4 ông Bạch Hưng Chế. Sinh 1 con trai là Bạch Hưng Hải và 3 con gái (Oanh, Ân, Yến)."
    },
    {
        "stt": 24, "ma": "C1-D6-08", "can": "Can 1", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nhỏ", "huy": "Bạch Quốc Hưng", "gioi_tinh": "Nam",
        "nam_sinh": "1945 (Ất Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Chế (Tam)", "me": "Nguyễn Thị Tứ",
        "phoi_ngau": "Đào Thị Nữ Trang (sinh 1951, quê Bình Giang, Hải Dương)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 5 ông Bạch Hưng Chế. Bộ đội chống Mỹ, cư trú Bình Giang, Hải Dương. Sinh 2 trai: Thái Nguyên, Quang Ngọc; 2 gái: Hoàng Giang, Thu Hà."
    },
    # Đời 7 Can 1
    {
        "stt": 25, "ma": "C1-D7-01", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tùng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1951 (Tân Mão)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cháu", "me": "Nguyễn Thị Giá",
        "phoi_ngau": "Nguyễn Thị Lam Trà (sinh 1961)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Cháu. Bộ đội chống Mỹ, HC Chiến công GP hạng 3. P. Vinh Tân, TP Vinh. Con: Hương Giang (1990), Kiên (1998)."
    },
    {
        "stt": 26, "ma": "C1-D7-02", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Hùng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1956 (Bính Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cháu", "me": "Nguyễn Thị Giá",
        "phoi_ngau": "Hồ Thị Sơn (sinh 1956)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Cháu. Cư trú TP.HCM. Con: Thùy Dương (1982), Thủy Tiên (1986-2011), Trà Hoa Hạ (1997), Đạt (2000)."
    },
    {
        "stt": 27, "ma": "C1-D7-03", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Dũng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1959 (Kỷ Hợi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cháu", "me": "Nguyễn Thị Giá",
        "phoi_ngau": "Phạm Thị Mười (sinh 1961)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Cháu. Thượng tá, Phó phòng CSGT Công an Nghệ An. Con: Hương Trà (1986), Hương Thảo (1988)."
    },
    {
        "stt": 28, "ma": "C1-D7-04", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thất", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cháu", "me": "Nguyễn Thị Giá",
        "phoi_ngau": "Phạm Thị Hồng Nhạn (sinh 1970)", "mo_phan": "",
        "ghi_chu": "Con trai út ông Bạch Hưng Cháu. Cư trú P. Cửa Nam, TP Vinh. Con: Bạch Hưng An (1997), Bạch Thị Hải Lê (1999)."
    },
    {
        "stt": 29, "ma": "C1-D7-05", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Quốc Việt", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhạ", "me": "Nguyễn Thị Chắt",
        "phoi_ngau": "Phạm Thị Tâm", "mo_phan": "",
        "ghi_chu": "Con trai trưởng ông Bạch Hưng Nhạ. Chiến đấu biên giới Tây Nam (1979). Con: Tuấn Dũng (1984), Hồng Duyên (1986), Mỹ Lệ (1990)."
    },
    {
        "stt": 30, "ma": "C1-D7-06", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Nam", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1960 (Canh Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhạ", "me": "Nguyễn Thị Chắt",
        "phoi_ngau": "Nguyễn Thị Lan (sinh 1963)", "mo_phan": "",
        "ghi_chu": "Con trai thứ ông Bạch Hưng Nhạ. Chủ tịch xã Nghĩa Đức, Nghĩa Đàn. Con: Bảo Trung (1986), Khánh Chi (1988), Thùy (1990)."
    },
    {
        "stt": 31, "ma": "C1-D7-07", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lập", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958", "nam_mat": "1968 (mất ngày 17/6/1968)",
        "ngay_gio": "", "cha": "Bạch Hưng Mận", "me": "Nguyễn Thị Loan",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai cả ông Bạch Hưng Mận, mất nhỏ."
    },
    {
        "stt": 32, "ma": "C1-D7-08", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thanh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1961 (Tân Sửu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mận", "me": "Nguyễn Thị Loan",
        "phoi_ngau": "Hoa Thị Huyền (sinh 1963)", "mo_phan": "",
        "ghi_chu": "Con trai thứ của ông Bạch Hưng Mận. Xã Nghĩa Đức, Nghĩa Đàn. Con: Hoàn (1987), Thương (1989-1994), Hảo (1991), Hạnh (1995), Phúc (1997)."
    },
    {
        "stt": 33, "ma": "C1-D7-09", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Diệu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1968 (Mậu Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mận", "me": "Nguyễn Thị Loan",
        "phoi_ngau": "Hồ Thị Hương (sinh 1976)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Mận. Xã Nghĩa Đức. Con: Tùng Dương (1996), Quang Sinh (2001)."
    },
    {
        "stt": 34, "ma": "C1-D7-10", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lợi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1972 (Nhâm Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mận", "me": "Nguyễn Thị Loan",
        "phoi_ngau": "Trần Thị Tư (sinh 1978)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 4 ông Bạch Hưng Mận. Con: Bạch Hưng Thắng (2001)."
    },
    {
        "stt": 35, "ma": "C1-D7-11", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Hải", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1981 (Tân Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Mai", "me": "Hồ Thị Tình",
        "phoi_ngau": "Phan Thị Hoa (sinh 1984)", "mo_phan": "",
        "ghi_chu": "Con trai độc nhất của ông Bạch Hưng Mai. Nghĩa Đức, Nghĩa Đàn. Con: Quốc Quân (2005), Quyên (2012)."
    },
    {
        "stt": 36, "ma": "C1-D7-12", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Thái Nguyên", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1974 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhỏ", "me": "Đào Thị Nữ Trang",
        "phoi_ngau": "Trần Thị Anh (sinh 1982)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Nhỏ. Cư trú Bình Giang, Hải Dương. Con: Khánh An (2010)."
    },
    {
        "stt": 37, "ma": "C1-D7-13", "can": "Can 1", "doi": "Đời 7",
        "ho_ten": "Bạch Quang Ngọc", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1979 (Kỷ Mùi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nhỏ", "me": "Đào Thị Nữ Trang",
        "phoi_ngau": "Bùi Thị Phương (sinh 1985)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Nhỏ. TP Vật tư Cty CP Thép Hoàng Long. Con: Đức Thịnh (2013)."
    },

    # =========================================================================
    # CAN 2: TỔ BẠCH HƯNG KHÂU (CHUẨN XÁC THEO VĂN BẢN TRANG 35 - 61)
    # =========================================================================
    {
        "stt": 38, "ma": "C2-D3-01", "can": "Can 2", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Khâu", "huy": "Húy Hoài", "gioi_tinh": "Nam",
        "nam_sinh": "1821 (Tân Tỵ)", "nam_mat": "1897 (Đinh Dậu) - Thọ 77",
        "ngay_gio": "02/08 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Bùi Thị Muồng (giỗ 07/08 ÂL, không con); Bà 2: Phạm Thị Hoạt (thọ 62, giỗ 28/08 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 2. Con trai thứ 2 Cụ Lưu. Bà hai sinh 5 trai (Thấu, Thưu, Xu, Phù, Câu) và 1 gái (Bạch Thị Hai)."
    },
    # Đời 4 Can 2
    {
        "stt": 39, "ma": "C2-D4-01", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thấu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi chưa có gia thất",
        "ngay_gio": "05/04 ÂL", "cha": "Bạch Hưng Khâu", "me": "Phạm Thị Hoạt",
        "phoi_ngau": "", "mo_phan": "Kẹt Đá xứ",
        "ghi_chu": "Con trai cả Cụ Khâu."
    },
    {
        "stt": 40, "ma": "C2-D4-02", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Thưu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "25/10 ÂL", "cha": "Bạch Hưng Khâu", "me": "Phạm Thị Hoạt",
        "phoi_ngau": "Hồ Thị Út (giỗ 19/03 ÂL, mộ Thồ Lô)", "mo_phan": "Thồ Lô xứ",
        "ghi_chu": "Con trai thứ 2 Cụ Khâu. Sinh 1 con trai: Bạch Sĩ Xum (Cố Hòe) và 1 con gái: Bạch Thị Đĩ (chết nhỏ)."
    },
    {
        "stt": 41, "ma": "C2-D4-03", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Xu", "huy": "Húy Mâu", "gioi_tinh": "Nam",
        "nam_sinh": "1854 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "01/06 ÂL", "cha": "Bạch Hưng Khâu", "me": "Phạm Thị Hoạt",
        "phoi_ngau": "Ngô Thị Chuẩn (giỗ 04/01 ÂL)", "mo_phan": "Thồ Lô xứ",
        "ghi_chu": "Con trai thứ 3 Cụ Khâu. Sinh 4 trai: Tam (Cố Chắt Mãi), Hợi, Yêm, Em; và 3 gái: Chắt, Tứ, Em."
    },
    {
        "stt": 42, "ma": "C2-D4-04", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Phù", "huy": "Húy Tư Lang", "gioi_tinh": "Nam",
        "nam_sinh": "1862 (Nhâm Tuất)", "nam_mat": "1912 (Nhâm Tý) - Thọ 51",
        "ngay_gio": "28/06 ÂL", "cha": "Bạch Hưng Khâu", "me": "Phạm Thị Hoạt",
        "phoi_ngau": "Không lập gia thất", "mo_phan": "Khe Xứ",
        "ghi_chu": "Con trai thứ 4 Cụ Khâu. Văn bản ghi rõ: Ông không lập gia thất, mất khi 51 tuổi."
    },
    {
        "stt": 43, "ma": "C2-D4-05", "can": "Can 2", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Câu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1864 (Giáp Tý)", "nam_mat": "1894 (Giáp Ngọ) - Thọ 31",
        "ngay_gio": "28/09 ÂL", "cha": "Bạch Hưng Khâu", "me": "Phạm Thị Hoạt",
        "phoi_ngau": "Ngô Thị Tư (giỗ 24/09 ÂL)", "mo_phan": "Núi Khe Xứ",
        "ghi_chu": "Con trai thứ 5 Cụ Khâu. Sinh 3 trai: Bạch Hưng Phòng (1), Bạch Hưng Dĩ (2), Bạch Hưng Em (3) (húy Con); và con gái Bạch Thị Đị."
    },
    # Đời 5 Can 2
    {
        "stt": 44, "ma": "C2-D5-01", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Sĩ Xum", "huy": "Cố Hòe", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "11/08 ÂL", "cha": "Bạch Hưng Thưu", "me": "Hồ Thị Út",
        "phoi_ngau": "Hồ Thị Nhiên (giỗ 15/03 ÂL)", "mo_phan": "Thồ Lô xứ",
        "ghi_chu": "Con trai độc nhất của Cụ Bạch Hưng Thưu. Con cái không rõ."
    },
    {
        "stt": 45, "ma": "C2-D5-02", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Tam", "huy": "Cố Chắt Mãi", "gioi_tinh": "Nam",
        "nam_sinh": "1891 (Tân Mão)", "nam_mat": "1950 (Canh Dần) - Thọ 60",
        "ngay_gio": "22/02 ÂL", "cha": "Bạch Hưng Xu", "me": "Ngô Thị Chuẩn",
        "phoi_ngau": "Nguyễn Thị Em (1890 - 1964, thọ 76, giỗ 30/11 ÂL, mộ Khe Xứ)", "mo_phan": "Núi Khe Xứ",
        "ghi_chu": "Con trai cả Cụ Xu. Tuần Kiểm của làng. Sinh 1 trai độc nhất là Liệt sỹ Bạch Hưng Đình và 5 gái (Tâm, Cháu, Nhỏ, Tìu, Ca)."
    },
    {
        "stt": 46, "ma": "C2-D5-03", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Hợi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất lúc 51 tuổi",
        "ngay_gio": "23/01 ÂL", "cha": "Bạch Hưng Xu", "me": "Ngô Thị Chuẩn",
        "phoi_ngau": "Nguyễn Thị Tư (thọ 83, giỗ 26/01 ÂL)", "mo_phan": "Nghĩa Quang, Nghĩa Đàn",
        "ghi_chu": "Con trai thứ 2 Cụ Xu. Sinh 4 trai: Nậy, Tứ, Niêm, Công; và 4 gái: Hai, Bơ, Tám, Thất."
    },
    {
        "stt": 47, "ma": "C2-D5-04", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Yêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất ở tuổi 49",
        "ngay_gio": "", "cha": "Bạch Hưng Xu", "me": "Ngô Thị Chuẩn",
        "phoi_ngau": "Nguyễn Thị Hai (mất tuổi 53)", "mo_phan": "Thồ Lô xứ",
        "ghi_chu": "Con trai thứ 3 Cụ Xu. Sinh 2 trai: Nậy, Vỹ; và 2 gái: Bát, Mơ."
    },
    {
        "stt": 48, "ma": "C2-D5-05", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Con thứ 4 Cụ Xu", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất lúc chưa có gia thất",
        "ngay_gio": "06/08 ÂL", "cha": "Bạch Hưng Xu", "me": "Ngô Thị Chuẩn",
        "phoi_ngau": "Không vợ", "mo_phan": "Thồ Lô xứ",
        "ghi_chu": "Con trai thứ 4 Cụ Xu. Con thứ ông Bạch Hưng Niêm thờ tự."
    },
    {
        "stt": 49, "ma": "C2-D5-06", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Phòng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất lúc chưa có gia thất",
        "ngay_gio": "19/05 ÂL", "cha": "Bạch Hưng Câu", "me": "Ngô Thị Tư",
        "phoi_ngau": "Không vợ", "mo_phan": "Núi Khe Xứ",
        "ghi_chu": "Con trai cả Cụ Câu. Ông Bạch Hưng Quyền thờ tự."
    },
    {
        "stt": 50, "ma": "C2-D5-07", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Dĩ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 61 tuổi",
        "ngay_gio": "17/03 ÂL", "cha": "Bạch Hưng Câu", "me": "Ngô Thị Tư",
        "phoi_ngau": "Nguyễn Thị Lù (giỗ 09/05 ÂL)", "mo_phan": "Núi Khe xứ (Rú sau)",
        "ghi_chu": "Con trai thứ 2 Cụ Câu. Sinh 3 con trai: Nậy (Dĩ), Cường, Bốn (Uyển); và 2 gái: Bô, Ngụ."
    },
    {
        "stt": 51, "ma": "C2-D5-08", "can": "Can 2", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Em", "huy": "Húy Con", "gioi_tinh": "Nam",
        "nam_sinh": "1900 (Canh Tý)", "nam_mat": "1981 - Thọ 82",
        "ngay_gio": "28/03 ÂL", "cha": "Bạch Hưng Câu", "me": "Ngô Thị Tư",
        "phoi_ngau": "Ngô Thị Chử (1901 - 1978, thọ 77, giỗ 15/12 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai thứ 3 Cụ Câu. Sinh 2 con trai: Liệt sỹ Bạch Hưng Lực, Liệt sỹ Bạch Hưng Xuân; và 2 con gái: Lân, La (Hồng)."
    },
    # Đời 6 Can 2
    {
        "stt": 52, "ma": "C2-D6-01", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Đình", "huy": "Liệt sỹ", "gioi_tinh": "Nam",
        "nam_sinh": "1925 (Ất Sửu)", "nam_mat": "1968 (Mậu Thân) - Hy sinh",
        "ngay_gio": "03/07 ÂL", "cha": "Bạch Hưng Tam", "me": "Nguyễn Thị Em",
        "phoi_ngau": "Bà 1: Hồ Thị Hảo (1922-1951, giỗ 30/01 ÂL); Bà 2: Nguyễn Thị Quế (sinh 1925, Hưng Châu)",
        "mo_phan": "Núi Khe Xứ / Bằng Tổ quốc ghi công",
        "ghi_chu": "Chủ tịch UBHC xã Hưng Phú, hy sinh 27/8/1968. Bà Hảo sinh: Thái, Phú. Bà Quế sinh: Dương, Tuyên, Trung, Bạch Thị Dung."
    },
    {
        "stt": 53, "ma": "C2-D6-02", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Con Cụ Hợi", "gioi_tinh": "Nam",
        "nam_sinh": "1918 (Mậu Ngọ)", "nam_mat": "1993 (Quý Dậu)",
        "ngay_gio": "13/06 ÂL", "cha": "Bạch Hưng Hợi", "me": "Nguyễn Thị Tư",
        "phoi_ngau": "Nguyễn Thị Mai (sinh 1927, Hà Tây)", "mo_phan": "Nghĩa Đức, Nghĩa Đàn",
        "ghi_chu": "Con trai cả Cụ Hợi. Sinh 3 trai: Vinh, Quang, Trung; và con gái Bạch Thị Vân."
    },
    {
        "stt": 54, "ma": "C2-D6-03", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Tứ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1923 (Quý Hợi)", "nam_mat": "1999 (Kỷ Mão)",
        "ngay_gio": "30/01 ÂL", "cha": "Bạch Hưng Hợi", "me": "Nguyễn Thị Tư",
        "phoi_ngau": "Phạm Thị Nhỏ (1923 - 1988, giỗ 23/09 ÂL)", "mo_phan": "Nghĩa Thịnh, Nghĩa Đàn",
        "ghi_chu": "Con trai thứ 2 Cụ Hợi. Sinh 2 trai: Trường, Thường; 3 gái: Thu, Hồng, Nga."
    },
    {
        "stt": 55, "ma": "C2-D6-04", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Niêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1930 (Canh Ngọ)", "nam_mat": "2011 (Tân Mão)",
        "ngay_gio": "11/05 ÂL", "cha": "Bạch Hưng Hợi", "me": "Nguyễn Thị Tư",
        "phoi_ngau": "Hồ Thị Huê (sinh 1930)", "mo_phan": "Hố Đót, Rú Trước",
        "ghi_chu": "Con trai thứ 3 Cụ Hợi. Sinh 3 trai: Thông, Tiến, Ty; 3 gái: Hoan, Ý, Vi."
    },
    {
        "stt": 56, "ma": "C2-D6-05", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Công", "huy": "Tên gọi: Ngọt", "gioi_tinh": "Nam",
        "nam_sinh": "1938 (Mậu Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hợi", "me": "Nguyễn Thị Tư",
        "phoi_ngau": "Nguyễn Thị Y (sinh 1940, Nghi Thủy, Cửa Lò)", "mo_phan": "",
        "ghi_chu": "Con trai út Cụ Hợi. Công nhân Nông trường 3.2, Minh Hợp, Quỳ Hợp. Sinh 4 trai: Tâm, Long, Phương, Kiều; 2 gái: Luyến, Khang."
    },
    {
        "stt": 57, "ma": "C2-D6-06", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Con Cụ Yêm", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất sớm",
        "ngay_gio": "", "cha": "Bạch Hưng Yêm", "me": "Nguyễn Thị Hai",
        "phoi_ngau": "", "mo_phan": "Sông Con",
        "ghi_chu": "Con cả Cụ Yêm. Đi bè ở Hương Sơn, sinh con trai là Bạch Hưng Chinh (ở Quỳ Hợp)."
    },
    {
        "stt": 58, "ma": "C2-D6-07", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Vỹ", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi chưa có con",
        "ngay_gio": "03/03 ÂL", "cha": "Bạch Hưng Yêm", "me": "Nguyễn Thị Hai",
        "phoi_ngau": "", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ 2 Cụ Yêm. Mắc bệnh suyễn, mất không con."
    },
    {
        "stt": 59, "ma": "C2-D6-08", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Nậy", "huy": "Con Cụ Dị", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khoảng 30 tuổi (trước 1950)",
        "ngay_gio": "11/08 ÂL", "cha": "Bạch Hưng Dĩ", "me": "Nguyễn Thị Lù",
        "phoi_ngau": "Nguyễn Thị Lục (sau cải giá Diễn Châu)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con cả Cụ Dĩ. Sinh 2 con gái: Tuyết (1945), Tý (Chiên, 1948). Chú Cường và Chú Bốn nuôi dưỡng."
    },
    {
        "stt": 60, "ma": "C2-D6-09", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Cường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1920 (Canh Thân)", "nam_mat": "2016 (Bính Thân) - Thọ 97",
        "ngay_gio": "24/05 ÂL", "cha": "Bạch Hưng Dĩ", "me": "Nguyễn Thị Lù",
        "phoi_ngau": "Nguyễn Thị Lam (1930 - 1968, giỗ 25/06 ÂL, mộ Khe xứ)", "mo_phan": "Đức Trọng, Lâm Đồng",
        "ghi_chu": "Cán bộ Tiền khởi nghĩa, 65 năm tuổi Đảng, HC Kháng chiến Hạng Nhất. Mất 24/05/2016 tại Lâm Đồng. Sinh 1 trai duy nhất là Bạch Hưng Huy và 7 gái (Thanh, Minh, Liên, Chiến, Thắng...)."
    },
    {
        "stt": 61, "ma": "C2-D6-10", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Bốn", "huy": "Ông Uyển", "gioi_tinh": "Nam",
        "nam_sinh": "1930 (Canh Ngọ)", "nam_mat": "1993 (Quý Dậu) - Thọ 64",
        "ngay_gio": "27/03 ÂL", "cha": "Bạch Hưng Dĩ", "me": "Nguyễn Thị Lù",
        "phoi_ngau": "Hồ Thị Tam (sinh 1932 Nhâm Thân, con ông Cựu Học)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Thẩm phán TAND huyện Hưng Nguyên, 30 năm tuổi Đảng, HC Kháng chiến Hạng Nhất. Sinh 6 con trai: Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn; 3 gái: Nga, Thành, Công."
    },
    {
        "stt": 62, "ma": "C2-D6-11", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Lực", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1931 (Tân Mùi)", "nam_mat": "2009 (Kỷ Sửu) - Thọ 79",
        "ngay_gio": "25/06 ÂL", "cha": "Bạch Hưng Em (Con)", "me": "Ngô Thị Chử",
        "phoi_ngau": "Phạm Thị Đích (1932 - 2014, thọ 83, giỗ 04/02 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai cả Cụ Em (Con). Cựu chiến binh Điện Biên Phủ, Thương binh chống Pháp, 40 năm tuổi Đảng. Sinh 6 trai: Mậu, Phượng, Loan (mất nhỏ), Trường, Sơn (mất nhỏ), Lường; 2 gái: Chắt, Châu."
    },
    {
        "stt": 63, "ma": "C2-D6-12", "can": "Can 2", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Xuân", "huy": "Liệt sỹ", "gioi_tinh": "Nam",
        "nam_sinh": "1932 (Nhâm Thân)", "nam_mat": "1966 (Bính Ngọ) - Hy sinh",
        "ngay_gio": "07/03 ÂL", "cha": "Bạch Hưng Em (Con)", "me": "Ngô Thị Chử",
        "phoi_ngau": "Hồ Thị Huyền (1935 - 2013, thọ 78, giỗ 20/11 ÂL)", "mo_phan": "Nghĩa trang QG Tuyên Hóa, Quảng Bình",
        "ghi_chu": "Liệt sỹ lái xe vận tải Trường Sơn (Bình Trị Thiên), hy sinh tháng 3/1966 tuổi 34. Sinh: Cầm, Quyền, Thư."
    },
    # Đời 7 Can 2
    {
        "stt": 64, "ma": "C2-D7-01", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thái", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1945 (Ất Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Hồ Thị Hảo",
        "phoi_ngau": "Phan Thị Hà (sinh 1952)", "mo_phan": "",
        "ghi_chu": "Con trai cả Liệt sỹ Đình. Bộ đội chống Mỹ, thương binh giải ngũ 1976. Xã Hưng Đạo. Con: Vân (1974), Bạch Hưng Thư (1978), Tuyết (1983), Thắm (1986)."
    },
    {
        "stt": 65, "ma": "C2-D7-02", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phú", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1949 (Kỷ Sửu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Hồ Thị Hảo",
        "phoi_ngau": "Nguyễn Thị Thủy (sinh 1953, Đức Thọ)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 Liệt sỹ Đình. Giáo viên cấp 3, TP Hà Tĩnh. Người phụng tự Tổ Can 4 Cụ Kế. Con: Bạch Hồng Y (1979), Bạch Ly Na (1983 - Thạc sĩ, BS Viện Nhi TW)."
    },
    {
        "stt": 66, "ma": "C2-D7-03", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Dương", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1954 (Giáp Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Quế",
        "phoi_ngau": "Lê Thị Anh (sinh 1956, Thầy thuốc ưu tú, BV Nhi Nghệ An)", "mo_phan": "",
        "ghi_chu": "Con thứ 3 Liệt sỹ Đình. Đại tá, Phó trưởng phòng An ninh Văn hóa Công an Nghệ An. Con: Bạch Hồng Hải (1987), Bạch Hưng Hoàn (1992)."
    },
    {
        "stt": 67, "ma": "C2-D7-04", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tuyên", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Quế",
        "phoi_ngau": "Nguyễn Thị Niêm (sinh 1960)", "mo_phan": "",
        "ghi_chu": "Con thứ 4 Liệt sỹ Đình. Giám đốc Công ty Thủy nông Nam Nghệ An. Con: Thùy Linh (1985), Bạch Hưng Tân (1987 - Thạc sĩ kinh tế)."
    },
    {
        "stt": 68, "ma": "C2-D7-05", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trung", "huy": "Con Liệt sỹ Đình", "gioi_tinh": "Nam",
        "nam_sinh": "1966 (Bính Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Đình", "me": "Nguyễn Thị Quế",
        "phoi_ngau": "Trần Thị Thúy Vân (sinh 1964)", "mo_phan": "",
        "ghi_chu": "Con trai út Liệt sỹ Đình. Trưởng trạm Ba Ra Nam Đàn. Con: Huyền An (1990), Bạch Hưng Ninh (1999)."
    },
    {
        "stt": 69, "ma": "C2-D7-06", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Vinh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1953 (Quý Tỵ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy", "me": "Nguyễn Thị Mai",
        "phoi_ngau": "Hoàng Thị Sâm (sinh 1964)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Nậy (Hợi). Con: Bạch Hưng Bảo (1976), Phương Diệu (1980), Bình (1990), Bạch Hưng Nguyên (1991)."
    },
    {
        "stt": 70, "ma": "C2-D7-07", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Quang", "huy": "", "gioi_tinh": "Nam", "nam_sinh": "1960", "nam_mat": "2003",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy", "me": "Nguyễn Thị Mai",
        "phoi_ngau": "Ngô Thị Hằng", "mo_phan": "Nghĩa Đức",
        "ghi_chu": "Con thứ 2 ông Bạch Hưng Nậy (Hợi). Con: Bạch Hưng Tuấn (1989), Bạch Hưng Hoàng (1993), Trinh (1994)."
    },
    {
        "stt": 71, "ma": "C2-D7-08", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trung", "huy": "Con ông Nậy (Hợi)", "gioi_tinh": "Nam",
        "nam_sinh": "1965 (Ất Tỵ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nậy", "me": "Nguyễn Thị Mai",
        "phoi_ngau": "Võ Thị Dung (sinh 1966)", "mo_phan": "",
        "ghi_chu": "Con thứ 3 ông Bạch Hưng Nậy (Hợi). Nghĩa Đức. Con: Tú (1989), Tài (1991), Tiến (1998)."
    },
    {
        "stt": 72, "ma": "C2-D7-09", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trường", "huy": "Con ông Tứ", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tứ", "me": "Phạm Thị Nhỏ",
        "phoi_ngau": "Lương Thị Yến (sinh 1965)", "mo_phan": "",
        "ghi_chu": "Con trai ông Bạch Hưng Tứ. Con: Trang (1988), Điểm (1990), Bạch Hưng Dũng (1994)."
    },
    {
        "stt": 73, "ma": "C2-D7-10", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1968 (Mậu Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tứ", "me": "Phạm Thị Nhỏ",
        "phoi_ngau": "Hồ Thị Hòa (sinh 1988)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Tứ. Con: Khánh (1993), Quỳnh (1995)."
    },
    {
        "stt": 74, "ma": "C2-D7-11", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thông", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1963 (Quý Mão)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Niêm", "me": "Hồ Thị Huê",
        "phoi_ngau": "Phạm Thị Lan (sinh 1968)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Niêm. Con: Bạch Hưng Dũng (1991), Bạch Hưng Sỹ (1993), Bạch Hưng Chung (2007)."
    },
    {
        "stt": 75, "ma": "C2-D7-12", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tiến", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1968 (Mậu Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Niêm", "me": "Hồ Thị Huê",
        "phoi_ngau": "Trần Thị Xoan (sinh 1981, Vũ Quang)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Niêm. Con: Bạch Hưng Quang Vũ (2009), Bảo Anh (2012)."
    },
    {
        "stt": 76, "ma": "C2-D7-13", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Ty", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1973 (Quý Sửu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Niêm", "me": "Hồ Thị Huê",
        "phoi_ngau": "Phạm Thị Đức", "mo_phan": "",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Niêm. Con: Bạch Hưng Thanh, Bạch Thị Trang."
    },
    {
        "stt": 77, "ma": "C2-D7-14", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Tâm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1967 (Đinh Mùi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Công", "me": "Nguyễn Thị Y",
        "phoi_ngau": "Tô Thị Phương (sinh 1969)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Công. Minh Hợp, Quỳ Hợp. Con: Bạch Hưng Hùng (1992), Vinh (1998)."
    },
    {
        "stt": 78, "ma": "C2-D7-15", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Long", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1970 (Canh Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Công", "me": "Nguyễn Thị Y",
        "phoi_ngau": "Nguyễn Thị Phương (sinh 1974)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Công. Minh Hợp, Quỳ Hợp. Con: Hiền (1993), Bạch Hưng Hậu (1999)."
    },
    {
        "stt": 79, "ma": "C2-D7-16", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phương", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1972 (Nhâm Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Công", "me": "Nguyễn Thị Y",
        "phoi_ngau": "Nguyễn Thị Thúy (sinh 1972)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Công. Minh Hợp, Quỳ Hợp. Con: Vân (1999), Hoa (2002)."
    },
    {
        "stt": 80, "ma": "C2-D7-17", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Kiều", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1976 (Bính Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Công", "me": "Nguyễn Thị Y",
        "phoi_ngau": "Trương Thị Lý (sinh 1977)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 4 ông Bạch Hưng Công. Nông trường 3.2. Con: Bạch Hưng Luân (2001), Giang (2008)."
    },
    {
        "stt": 81, "ma": "C2-D7-18", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Huy", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1952 (Nhâm Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Cường", "me": "Nguyễn Thị Lam",
        "phoi_ngau": "Võ Thị Chiên (1962 - 2012, giỗ 01/04 ÂL)", "mo_phan": "Đức Trọng, Lâm Đồng",
        "ghi_chu": "Con trai độc nhất của Cụ Bạch Hưng Cường. Giáo viên cấp 3, cư trú TT Liên Nghĩa, Đức Trọng, Lâm Đồng. Con: Hường (1988), Hà (1990), Quỳnh (2000)."
    },
    {
        "stt": 82, "ma": "C2-D7-19", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Uyển", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1952 (Nhâm Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "Ngô Thị Thái (1954 - mất 06/01/2016 Bính Thân, tức 13/02/2016 DL)", "mo_phan": "",
        "ghi_chu": "Con trai cả ông Bạch Hưng Bốn. Thượng tá Quân đội, Ban biên soạn Gia phả 2014. Con: Bạch Hưng Trinh (1980 - Thạc sĩ CNTT), Bạch Hưng Tường (1984 - KS Xây dựng)."
    },
    {
        "stt": 83, "ma": "C2-D7-20", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Triển", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1956 (Bính Thân)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "Nguyễn Thị Nhung (sinh 1958, Thiếu tá QĐ)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 ông Bạch Hưng Bốn. Đại tá, Phó phòng Chính trị Công an Nghệ An. Con: Thảo Nguyên (1984), Bạch Quốc Tuyên (1986 - Thạc sĩ, Bí thư Đoàn TN Học viện ANND)."
    },
    {
        "stt": 84, "ma": "C2-D7-21", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lịch", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1961 (Tân Sửu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "Trần Thị Chang (sinh 1964, Phú Thọ)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 3 ông Bạch Hưng Bốn. Thượng úy Công an, Cục Đăng kiểm VN. Con: Bạch Hưng Kiên (1990), Mỹ Hạnh (1995)."
    },
    {
        "stt": 85, "ma": "C2-D7-22", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thịnh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1963 (Quý Mão)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "Phạm Thị Xuân (sinh 1968)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 4 ông Bạch Hưng Bốn. Thạc sĩ, Đại tá Học viện Phòng không - Không quân QĐNDVN. Con: Thu Thảo (1995), Bạch Đình Khang (2002)."
    },
    {
        "stt": 86, "ma": "C2-D7-23", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Vĩnh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1971 (Tân Hợi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "Phạm Thị Hạnh (sinh 1976)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 5 ông Bạch Hưng Bốn. Cty Thủy nông Nam Nghệ An. Con: Hà Ngân (1999), Bạch Hưng Thủy (2006)."
    },
    {
        "stt": 87, "ma": "C2-D7-24", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Viễn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1975 (Ất Mão)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Bốn (Uyển)", "me": "Hồ Thị Tam",
        "phoi_ngau": "", "mo_phan": "", "ghi_chu": "Con trai thứ 6 ông Bạch Hưng Bốn."
    },
    {
        "stt": 88, "ma": "C2-D7-25", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Mậu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1958 (Mậu Tuất)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Phạm Thị Đích",
        "phoi_ngau": "Trần Thị Ca (sinh 1956)", "mo_phan": "",
        "ghi_chu": "Con trai đầu của ông Bạch Hưng Lực. Thượng úy pháo binh. Con: Bạch Hưng Tài (1984), Bạch Hưng Toản (1985)."
    },
    {
        "stt": 89, "ma": "C2-D7-26", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Phượng", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Phạm Thị Đích",
        "phoi_ngau": "Nguyễn Thị Trầm (Đức Thọ)", "mo_phan": "",
        "ghi_chu": "Con trai thứ của ông Bạch Hưng Lực. P. Trường Thi, TP Vinh. Con: Trang (1988), Bạch Hưng Đạt (1995)."
    },
    {
        "stt": 90, "ma": "C2-D7-27", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Trường", "huy": "Con ông Lực", "gioi_tinh": "Nam",
        "nam_sinh": "1969 (Kỷ Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Phạm Thị Đích",
        "phoi_ngau": "Ngô Thị Hằng (sinh 1971)", "mo_phan": "",
        "ghi_chu": "Con trai thứ của ông Bạch Hưng Lực. P. Trường Thi, TP Vinh. Con: Bạch Hưng Mạnh (1996), Bạch Hưng Khánh (2007)."
    },
    {
        "stt": 91, "ma": "C2-D7-28", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Lường", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1972 (Nhâm Tý)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lực", "me": "Phạm Thị Đích",
        "phoi_ngau": "Nguyễn Thị Hoài (sinh 1976)", "mo_phan": "",
        "ghi_chu": "Con trai út của ông Bạch Hưng Lực. Con: Bạch Hưng Kiên (1998), Bạch Hưng Dũng (2011)."
    },
    {
        "stt": 92, "ma": "C2-D7-29", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Cầm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1962 (Nhâm Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xuân", "me": "Hồ Thị Huyền",
        "phoi_ngau": "Nguyễn Thị Hà (sinh 1966)", "mo_phan": "",
        "ghi_chu": "Con trai đầu Liệt sỹ Xuân. Trưởng phòng KHVT Cty CP Cơ khí Vinh, P. Hà Huy Tập, TP Vinh. Con: Hoài Anh (1990), Trà My (1994), Thúy Hiền (1997), Bạch Bình Minh (2004 - HC Đồng Olympic Tiếng Anh Quốc gia 2015)."
    },
    {
        "stt": 93, "ma": "C2-D7-30", "can": "Can 2", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Quyền", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1964 (Giáp Thìn)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Xuân", "me": "Hồ Thị Huyền",
        "phoi_ngau": "Nguyễn Thị Tuyết (sinh 1969, quê Hưng Đạo, XKLĐ Anh)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 Liệt sỹ Xuân. Bảo vệ Làng SOS Vinh, cư trú P. Hà Huy Tập, TP Vinh. Con: Bạch Thị Liên (1993, SV HV Bưu chính), Bạch Hưng Khang (1996)."
    },

    # =========================================================================
    # CAN 3: TỔ BẠCH HƯNG NGHI (CHUẨN XÁC THEO VĂN BẢN TRANG 64 - 74)
    # =========================================================================
    {
        "stt": 94, "ma": "C3-D3-01", "can": "Can 3", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Nghi", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1829 (Kỷ Sửu)", "nam_mat": "Thọ 89 tuổi",
        "ngay_gio": "10/04 ÂL", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Bà 1: Hồ Thị Phong (thọ 81, giỗ 14/11 ÂL, không con); Bà 2: Nguyễn Thị Ngụ (húy Thông, thọ 36, giỗ 30/06 ÂL)",
        "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Khởi tổ Can 3. Thủ chỉ xã Phú Điền. Bà hai sinh: Bạch Thị Tâm, Bạch Hưng Truyện."
    },
    {
        "stt": 95, "ma": "C3-D4-01", "can": "Can 3", "doi": "Đời 4",
        "ho_ten": "Bạch Hưng Truyện", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất năm 28 tuổi",
        "ngay_gio": "11/07 ÂL", "cha": "Bạch Hưng Nghi", "me": "Nguyễn Thị Ngụ",
        "phoi_ngau": "Nguyễn Thị Năm (người Mệ Giụ, sau cải giá)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai độc nhất Cụ Nghi. 23 tuổi làm Lý trưởng, 28 tuổi bị Cai tổng hại. Sinh 2 con trai: Bạch Hưng Nguyễn, Bạch Hưng Khuyến."
    },
    {
        "stt": 96, "ma": "C3-D5-01", "can": "Can 3", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Nguyễn", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "",
        "ngay_gio": "30/12 ÂL", "cha": "Bạch Hưng Truyện", "me": "Nguyễn Thị Năm",
        "phoi_ngau": "Ngô Thị Đôi (thọ 66, giỗ 15/02 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai trưởng Cụ Truyện. Sinh 4 trai: Hai, Tư, Niêm, Sáu; 3 gái: Chắt, Tam, Nhiu."
    },
    {
        "stt": 97, "ma": "C3-D5-02", "can": "Can 3", "doi": "Đời 5",
        "ho_ten": "Bạch Hưng Khuyến", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 66 tuổi",
        "ngay_gio": "06/09 ÂL", "cha": "Bạch Hưng Truyện", "me": "Nguyễn Thị Năm",
        "phoi_ngau": "Hồ Thị Em (thọ 73, giỗ 26/07 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ Cụ Truyện. Cai xã Phú Điền. Sinh 2 trai: Hiến, Hưng; 4 gái: Đị, Em, Bưởi, Cọt."
    },
    # Đời 6 Can 3
    {
        "stt": 98, "ma": "C3-D6-01", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hai", "huy": "Kiểm Nhưng", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Thọ 54 tuổi",
        "ngay_gio": "20/08 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Phạm Thị Đị (giỗ 15/04 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai trưởng Cụ Nguyễn. Hương Kiểm của làng. Sinh Bạch Hưng Thành và 2 con gái (Hoàn, Sâm)."
    },
    {
        "stt": 99, "ma": "C3-D6-02", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Tư", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1906 (Bính Ngọ)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Phạm Thị Hai (sinh 1914)", "mo_phan": "",
        "ghi_chu": "Con trai thứ 2 Cụ Nguyễn. Sinh Bạch Hưng Liên (1943) và 3 gái (Chắt, Hồng, Mạo)."
    },
    {
        "stt": 100, "ma": "C3-D6-03", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Niêm", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất khi chưa lấy vợ",
        "ngay_gio": "15/05 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ 3 Cụ Nguyễn. Ông Bạch Hưng Liên phụng tự."
    },
    {
        "stt": 101, "ma": "C3-D6-04", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Sáu", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất ở tuổi 35",
        "ngay_gio": "28/06 ÂL", "cha": "Bạch Hưng Nguyễn", "me": "Ngô Thị Đôi",
        "phoi_ngau": "Vợ đi cải giá", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ 4 Cụ Nguyễn. Mất khi chưa có con."
    },
    {
        "stt": 102, "ma": "C3-D6-05", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hiến", "huy": "Đầu phiên Hiến", "gioi_tinh": "Nam",
        "nam_sinh": "1906 (Bính Ngọ)", "nam_mat": "1986 (Bính Dần) - Thọ 82",
        "ngay_gio": "15/02 ÂL", "cha": "Bạch Hưng Khuyến", "me": "Hồ Thị Em",
        "phoi_ngau": "Cao Thị Hoét (thọ 40, giỗ 01/08 ÂL)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con trai cả Cụ Khuyến. Giữ chức Đầu phiên của làng. Sinh 2 trai: Văn, Minh; và con gái Bạch Thị Lan."
    },
    {
        "stt": 103, "ma": "C3-D6-06", "can": "Can 3", "doi": "Đời 6",
        "ho_ten": "Bạch Hưng Hưng", "huy": "Húy Kiến", "gioi_tinh": "Nam",
        "nam_sinh": "", "nam_mat": "Mất lúc 14 tuổi",
        "ngay_gio": "12/03 ÂL", "cha": "Bạch Hưng Khuyến", "me": "Hồ Thị Em",
        "phoi_ngau": "", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Con thứ Cụ Khuyến. Con ông Hiến thờ tự."
    },
    # Đời 7 Can 3
    {
        "stt": 104, "ma": "C3-D7-01", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Thành", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1936 (Bính Tý)", "nam_mat": "Mất ở tuổi 49",
        "ngay_gio": "12/07 ÂL", "cha": "Bạch Hưng Hai (Kiểm Nhưng)", "me": "Phạm Thị Đị",
        "phoi_ngau": "Hồ Thị Tý (sinh 1936)", "mo_phan": "Núi Khe xứ",
        "ghi_chu": "Bí thư Đoàn xã, Cán bộ MTTQ Huyện, Cán bộ Ban GT 470, Huân chương KC Hạng Nhất. Sinh 5 trai: Chiến, Sỹ, Hương, Thủy, Triều; 4 gái: Hòa, Nữ, Trinh, Tâm."
    },
    {
        "stt": 105, "ma": "C3-D7-02", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Liên", "huy": "Thương binh 1/4", "gioi_tinh": "Nam",
        "nam_sinh": "1943 (Quý Mùi)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Tư", "me": "Phạm Thị Hai",
        "phoi_ngau": "Hồ Thị Miêng (sinh 1941)", "mo_phan": "",
        "ghi_chu": "Nhập ngũ 1964, Đường 9 Nam Lào; lái xe, bị thương hỏng cả 2 mắt. Thương binh nặng 1/4. Huân chương Kháng chiến Hạng 2. Con: Thắng (1970), Đức (1973 - mất trẻ), Lài (1979)."
    },
    {
        "stt": 106, "ma": "C3-D7-03", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Văn", "huy": "Thương binh 2/4", "gioi_tinh": "Nam",
        "nam_sinh": "1945 (Ất Dậu)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Hiến", "me": "Cao Thị Hoét",
        "phoi_ngau": "Trần Thị Kiều (sinh 1948, Thái Bình)", "mo_phan": "",
        "ghi_chu": "Pháo cao xạ Bắc Lào, Đường 9, Đường 20; bị thương cụt bàn tay trái. Thương binh 2/4. Huân chương Kháng chiến Hạng 2. Ban biên soạn Gia phả 2014. Con: Giang (1971), Tình (1979 - GV Toán), Yến (1973), Oanh (1975), Nhung (1978)."
    },
    {
        "stt": 107, "ma": "C3-D7-04", "can": "Can 3", "doi": "Đời 7",
        "ho_ten": "Bạch Hưng Minh", "huy": "", "gioi_tinh": "Nam",
        "nam_sinh": "1948 (Mậu Tý)", "nam_mat": "1993 - Thọ 47",
        "ngay_gio": "04/03 ÂL", "cha": "Bạch Hưng Hiến", "me": "Cao Thị Hoét",
        "phoi_ngau": "Võ Thị Xuân (Hưng Khánh)", "mo_phan": "",
        "ghi_chu": "Bộ đội Trung đoàn 270 tại Quảng Trị, Kế toán trưởng Cty Vật tư NN Nghệ An, cử nhân Kinh tế, Đảng viên 1967. HC Chiến công, HC Kháng chiến Hạng 3. Con: Sơn (1976), Hải (1977 - Thượng úy QK4), Hiền (1987)."
    },

    # =========================================================================
    # CAN 4: TỔ BẠCH HƯNG KẾ (VĂN BẢN TRANG 75)
    # =========================================================================
    {
        "stt": 108, "ma": "C4-D3-01", "can": "Can 4", "doi": "Đời 3",
        "ho_ten": "Bạch Hưng Kế", "huy": "Húy Cơ", "gioi_tinh": "Nam",
        "nam_sinh": "1854 (Giáp Dần)", "nam_mat": "",
        "ngay_gio": "", "cha": "Bạch Hưng Lưu", "me": "Phạm Thị Biềng",
        "phoi_ngau": "Ngô Thị Tam (húy Chư)", "mo_phan": "Xứ Thồ Lô",
        "ghi_chu": "Khởi tổ Can 4. Con trai thứ 4 Cụ Lưu. Con trai Bạch Hưng Thế mất năm 13 tuổi, 2 gái Mây và Út. Vô tự, được chi họ phụng tự. Trước đây ông Bạch Hưng Đình đảm nhận; nay con trai thứ 2 là ông Bạch Hưng Phú tiếp nối thờ tự."
    }
]

# GHI FILE EXCEL
print("Đang xuất bản Excel chuẩn xác theo văn bản...")
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
    ("Cha", 22, align_left),
    ("Mẹ", 20, align_left),
    ("Phối ngẫu", 32, align_left),
    ("Mộ phần", 22, align_left),
    ("Công trạng / Ghi chú (Chuẩn theo văn bản)", 45, align_left),
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

if __name__ == "__main__":
    # Sheet 1: Toàn bộ nhân đinh
    ws_all = wb.active
    ws_all.title = "Toàn bộ nhân đinh"
    setup_ws(ws_all, "DANH BẠ NHÂN ĐINH DÒNG HỌ BẠCH (CHUẨN XÁC THEO VĂN BẢN)")
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

    wb.save(EXCEL_PATH)
    print("Đã cập nhật xong file Excel thành công!")
