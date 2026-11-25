# -*- coding: utf-8 -*-
"""
Script xuất bản Gia_Pha_Ho_Bach.md và Gia_Pha_Viewer.html
Dựa 100% vào văn bản (không dùng sơ đồ).
"""
import os
import json
from update_excel import records

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

OUTPUT_DIR = r"d:\python\gia pha ho Bach"
MD_PATH = os.path.join(OUTPUT_DIR, "Gia_Pha_Ho_Bach.md")
HTML_PATH = os.path.join(OUTPUT_DIR, "Gia_Pha_Viewer.html")

# ==============================================================================
# 1. TẠO GIA_PHA_HO_BACH.MD
# ==============================================================================
md_content = """# GIA PHẢ HỌ BẠCH
## TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2 - THUỘC BẠCH TỘC ĐẠI TÔN
**Địa danh:** Làng Phú Điền, Xã Hưng Phú, Huyện Hưng Nguyên, Tỉnh Nghệ An  
**Thời gian biên soạn:** Năm Giáp Ngọ (2014) - Số hóa & Hiệu đính chuẩn xác: Năm 2026  
**Nguyên tắc số hóa:** **Phân tích quan hệ huyết thống 100% DỰA VÀO VĂN BẢN (TEXT), LOẠI BỎ SAI LỆCH DO SƠ ĐỒ VẼ GỘP CỘT NĂM 2014.**

---

> ### ⚠️ LƯU Ý ĐẶC BIỆT VỀ QUY TẮC PHÂN TÍCH QUAN HỆ HUYẾT THỐNG
> Trong bản in sách gia phả năm 2014, các trang sơ đồ vẽ hình khối (như trang 10-15, 34, 63) do người vẽ dồn cột, gióng đường kẻ giản lược nên đã xảy ra sự sai lệch nghiêm trọng về quan hệ cha con giữa các nhánh (ví dụ: vẽ gộp con Cụ Câu vào Cụ Phù; gộp con Cụ Đình vào Cụ Nậy; gộp con Cụ Bốn vào Cụ Lực; gộp con Cụ Công vào Cụ Cường...).  
> **Toàn bộ nội dung trong bản gia phả số hóa này được phục dựng và đối chiếu chính xác 100% DỰA THEO VĂN BẢN MÔ TẢ ĐÍCH DANH Ở TỪNG TRANG SÁCH**, đảm bảo tính chân thực lịch sử tuyệt đối cho con cháu muôn đời sau.

---

## MỤC LỤC
1. [Lời Nói Đầu & Cội Nguồn Dòng Tộc](#1-lời-nói-đầu--cội-nguồn-dòng-tộc)
2. [Vài Nét Về Lịch Sử Họ Bạch & Danh Nhân Bạch Liêu](#2-vài-nét-về-lịch-sử-họ-bạch--danh-nhân-bạch-liêu)
3. [Tổ Khởi Phát Trung Chi 2: Cụ Bạch Hưng Hóa (1730 - 1802)](#3-tổ-khởi-phát-trung-chi-2-cụ-bạch-hưng-hóa-1730---1802)
4. [Sơ Đồ Phả Hệ Tổng Thể](#4-sơ-đồ-phả-hệ-tổng-thể)
5. [Đời Thứ Nhất: Cụ Bạch Trần Tạo (1772 - 1801)](#5-đời-thứ-nhất-cụ-bạch-trần-tạo-1772---1801)
6. [Đời Thứ Hai: Cụ Bạch Hưng Khâm & Cụ Bạch Hưng Lưu](#6-đời-thứ-hai-cụ-bạch-hưng-khâm--cụ-bạch-hưng-lưu)
7. [CAN THỨ NHẤT: TỔ CAN BẠCH HƯNG CƯ (Đời 3 đến Đời 8)](#7-can-thứ-nhất-tổ-can-bạch-hưng-cư-đời-3-đến-đời-8)
8. [CAN THỨ HAI: TỔ CAN BẠCH HƯNG KHÂU (Đời 3 đến Đời 8)](#8-can-thứ-hai-tổ-can-bạch-hưng-khâu-đời-3-đến-đời-8)
9. [CAN THỨ BA: TỔ CAN BẠCH HƯNG NGHI (Đời 3 đến Đời 8)](#9-can-thứ-ba-tổ-can-bạch-hưng-nghi-đời-3-đến-đời-8)
10. [CAN THỨ TƯ: TỔ CAN BẠCH HƯNG KẾ (Húy Cơ, 1854)](#10-can-thứ-tư-tổ-can-bạch-hưng-kế-húy-cơ-1854)
11. [Bảng Tra Cứu Toàn Bộ 82 Ngày Kỵ Giỗ Trong Năm (Âm Lịch)](#11-bảng-tra-cứu-toàn-bộ-82-ngày-kỵ-giỗ-trong-năm-âm-lịch)

---

## 1. LỜI NÓI ĐẦU & CỘI NGUỒN DÒNG TỘC

*\"Con người ta sinh ra ở đời ai cũng có Tổ, có Tông, như cây có cội, như sông có nguồn. Đó là đạo lý truyền thống ngàn đời của dân tộc Việt Nam ta.\"*

Cuốn phả hệ này là **Tôn Đồ Tự Phả của Tiểu Chi 1, Trung Chi 2 thuộc Bạch Tộc Đại Tôn**, nguyên quán tại làng Phú Điền, xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An.

Vào những năm 1990, ông **Bạch Hưng Thủng** đã có công đầu tìm tòi, sưu tầm từ các bản gia phả bằng chữ Hán cổ của dòng họ, dịch thuật sang chữ Quốc ngữ để lưu truyền cho con cháu.

Đến năm 2014 (Giáp Ngọ), nhận thấy nhu cầu cập nhật các thế hệ con cháu nối tiếp, đồng thời để bảo tồn nguyên vẹn các di huấn, mộ phần và ngày kỵ giỗ của các bậc tiền nhân, ban tộc biểu gồm các ông:
- **Bạch Hưng Thái**
- **Bạch Hưng Phú**
- **Bạch Hưng Văn**
- **Bạch Hưng Uyển**
- **Bạch Hưng Tùng**

đã cùng nhau đối chiếu, bổ sung, chỉnh lý và in ấn thành tập sách gia phả hoàn chỉnh này nhằm giúp các thế hệ con cháu muôn đời sau:
1. **Biết rõ cội nguồn, gốc tích và mối quan hệ huyết thống chính xác giữa các nhánh, các Can.**
2. **Nắm vững ngày giỗ Âm lịch, nơi an táng (mộ phần) của tổ tiên để đời đời phụng tự, khói hương chu đáo.**
3. **Noi gương truyền thống hiếu học, yêu nước, kiên trung của các bậc tiền bối, nỗ lực phấn đấu làm rạng danh dòng họ Bạch.**

---

## 2. VÀI NÉT VỀ LỊCH SỬ HỌ BẠCH & DANH NHÂN BẠCH LIÊU

### 1. Thủy tổ Trạng nguyên Bạch Liêu (1236 - 1315)
- Dòng họ Bạch Việt Nam tự hào có cụ Thủy tổ là **Trạng nguyên Bạch Liêu**, sinh ngày 15 tháng Giêng năm Bính Thìn (1236) tại làng Yên Xá, huyện Chân Phúc (nay thuộc xã Mã Thành, huyện Yên Thành, tỉnh Nghệ An).
- Năm Bính Dần (1266), niên hiệu Thiệu Long thứ 9 đời vua Trần Thánh Tông, ngài đỗ **Đệ nhất giáp tiến sĩ đệ nhất danh (Trạng nguyên)** khi tròn 31 tuổi.
- Ngài nổi tiếng là bậc thông minh xuất chúng, học rộng tài cao, nhớ lâu và có tài năng kiệt xuất về chính trị, quân sự và bang giao. Sau khi đỗ Trạng nguyên, Thượng phụ Chiêu Minh Đại vương Trần Quang Khải rất mến phục đức tài của ngài, đã mời ngài về làm Môn khách trong phủ.
- Trong cuộc kháng chiến chống quân xâm lược Nguyên - Mông lần thứ 2 (1285) và lần thứ 3 (1287 - 1288), ngài có công lao to lớn giúp Trần Quang Khải hoạch định cơ mưu chiến lược, thảo hịch đánh giặc, góp phần vào những chiến thắng hiển hách như Chương Dương, Hàm Tử.
- Ngài cũng là tác giả của **\"Biến pháp tam chương\"** (Ba chương phép biến đổi), dâng lên vua Trần nhằm cải cách hành chính, quản lý đinh số, ruộng đất và củng cố quân đội, đặt nền móng vững chắc cho nền thịnh trị của nhà Trần.
- Cụ tạ thế ngày 24 tháng Giêng năm Ất Mão (1315), thọ 80 tuổi. Triều đình truy phong tước hiệu cao quý, nhân dân nhiều nơi lập đền thờ phụng để ghi nhớ công đức.

### 2. Đền thờ họ Bạch tại Hưng Phú, Hưng Nguyên
- Tại xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An, con cháu họ Bạch đã lập đền thờ Cụ Trạng nguyên Bạch Liêu và các bậc Tiên tổ.
- Đền thờ dòng họ Bạch tại Hưng Phú đã được Nhà nước xếp hạng là **Di tích Lịch sử - Văn hóa cấp Tỉnh** theo Quyết định số 2015/QĐ-UBND ngày 16/12/1993.
- Đến năm 2013, đền thờ đã được toàn thể con cháu dòng họ chung sức tôn tạo, nâng cấp khang trang, bề thế, trở thành nơi quy tụ tâm linh thiêng liêng của bà con cô bác họ Bạch khắp mọi miền đất nước.

---

## 3. TỔ KHỞI PHÁT TRUNG CHI 2: CỤ BẠCH HƯNG HÓA (1730 - 1802)

- **Năm sinh:** Canh Tuất (1730).
- **Năm mất:** Nhâm Tuất (1802), hưởng thọ 73 tuổi.
- **Ngày giỗ:** Ngày 14 tháng 4 Âm lịch.
- **Nơi an táng:** Mộ táng tại xứ Cồn Xanh.
- **Phối ngẫu (Các cụ bà):**
  1. **Bà thứ nhất:** Cụ bà **Bùi Thị Hàng Nhất**, ngày giỗ mồng 7 tháng 8 Âm lịch, mộ an táng tại xứ Cồn Xanh.
  2. **Bà thứ hai:** Cụ bà **Bùi Thị Hàng Hai**, ngày giỗ 28 tháng Giêng Âm lịch, mộ an táng tại xứ Khe Dài.
  3. **Bà thứ ba:** Cụ bà **Nguyễn Thị Hàng Hai**, ngày giỗ 11 tháng 8 Âm lịch, mộ an táng tại xứ Khe Dài.
- **Hậu duệ (8 người con):**
  1. **Bạch Trần Tạo** (con trai trưởng) -> Lập nên **Tiểu Chi 1** (chính là chi phái của cuốn phả hệ này).
  2. **Bạch Trần Điển** (con trai thứ 2) -> Lập nên **Tiểu Chi 2**.
  3. **Bạch Trần Khơm** (con trai thứ 3) -> Lập nên **Tiểu Chi 3**.
  4. **Bạch Khắc Chẩn** (con trai thứ 4).
  5. **Bạch Hưng Đại** (con trai thứ 5).
  6. **Bạch Thị Thao** (con gái).
  7. **Bạch Thị Đỏ** (con gái).
  8. (Một người con gái khác).

---

## 4. SƠ ĐỒ PHẢ HỆ TỔNG THỂ

```mermaid
graph TD
    BHH["TỔ KHỞI PHÁT TRUNG CHI 2<br/><b>BẠCH HƯNG HÓA</b> (1730 - 1802)<br/>Giỗ: 14/4 AL | Mộ: Cồn Xanh"]
    
    BTT["ĐỜI 1 - TIỂU CHI 1<br/><b>BẠCH TRẦN TẠO</b> (1772 - 1801)<br/>Giỗ: 28/9 AL | Vợ: Phạm Thị Hàng Nhất"]
    TC2["Tiểu Chi 2<br/>Bạch Trần Điển"]
    TC3["Tiểu Chi 3<br/>Bạch Trần Khơm"]
    
    BHH --> BTT
    BHH --> TC2
    BHH --> TC3
    
    BHKham["ĐỜI 2<br/>Bạch Hưng Khâm<br/>(Mất năm 20 tuổi)"]
    BHLuu["ĐỜI 2<br/><b>BẠCH HƯNG LƯU</b> (1798 - 1873)<br/>Húy Túc | Thọ 76 | Giỗ: 29/2 AL<br/>Vợ: Phạm Thị Biềng"]
    
    BTT --> BHKham
    BTT --> BHLuu
    
    C1["CAN 1<br/><b>BẠCH HƯNG CƯ</b><br/>(Húy Khuơ, 1819)<br/>Giỗ: 9/10 AL"]
    C2["CAN 2<br/><b>BẠCH HƯNG KHÂU</b><br/>(Húy Hoài, 1821-1897)<br/>Giỗ: 2/8 AL"]
    C3["CAN 3<br/><b>BẠCH HƯNG NGHI</b><br/>(1829, thọ 89)<br/>Giỗ: 10/4 AL"]
    C4["CAN 4<br/><b>BẠCH HƯNG KẾ</b><br/>(Húy Cơ, 1854)<br/>Vô tự - Phú tự"]
    
    BHLuu --> C1
    BHLuu --> C2
    BHLuu --> C3
    BHLuu --> C4
```

---

## 5. ĐỜI THỨ NHẤT: CỤ BẠCH TRẦN TẠO (1772 - 1801)

- **Họ và tên:** **BẠCH TRẦN TẠO** (Khởi tổ Đời 1 Tiểu Chi 1 - Trung Chi 2).
- **Sinh năm:** Nhâm Thìn (1772).
- **Mất năm:** Tân Dậu (1801), hưởng thọ 30 tuổi.
- **Ngày giỗ:** Ngày 28 tháng 9 Âm lịch.
- **Nơi an táng:** Mộ đặt tại xứ Tuần Tượng.
- **Phối ngẫu:** Cụ bà **Phạm Thị Hàng Nhất**, hưởng thọ 76 tuổi. Ngày giỗ mồng 8 tháng 10 Âm lịch. Mộ táng tại xứ Cồn Xanh.
- **Hậu duệ:** Cụ sinh hạ được 2 người con trai:
  1. **Bạch Hưng Khâm** (con trai trưởng).
  2. **Bạch Hưng Lưu** (con trai thứ 2).

---

## 6. ĐỜI THỨ HAI: CỤ BẠCH HƯNG KHÂM & CỤ BẠCH HƯNG LƯU

### 1. Cụ Bạch Hưng Khâm
- Là con trai trưởng của Cụ Bạch Trần Tạo.
- Cụ mất sớm năm 20 tuổi, chưa lập gia đình.
- **Ngày giỗ:** Ngày 28 tháng 5 Âm lịch.
- **Nơi an táng:** Mộ đặt tại xứ Cồn Vải.
- Con cháu nhánh Cụ Bạch Hưng Lưu đời đời phụng tự khói hương chu đáo.

### 2. Cụ Bạch Hưng Lưu (Húy Túc)
- Là con trai thứ 2 của Cụ Bạch Trần Tạo.
- **Sinh năm:** Mậu Ngọ (1798).
- **Mất năm:** Quý Dậu (1873), hưởng thọ 76 tuổi.
- **Ngày giỗ:** Ngày 29 tháng 2 Âm lịch (năm thiếu cúng ngày 28/2).
- **Nơi an táng:** Mộ táng tại xứ Thượng Phù.
- **Phối ngẫu:** Cụ bà **Phạm Thị Biềng**. Ngày giỗ mồng 10 tháng Giêng Âm lịch. Mộ táng tại xứ Thượng Phù.
- **Hậu duệ:** Cụ ông và cụ bà sinh được 4 người con trai (chia làm 4 Can) và 4 người con gái:
  - **4 người con trai (phân lập thành 4 Can):**
    1. **Bạch Hưng Cư** -> Lập nên **Can 1**.
    2. **Bạch Hưng Khâu** -> Lập nên **Can 2**.
    3. **Bạch Hưng Nghi** -> Lập nên **Can 3**.
    4. **Bạch Hưng Kế** -> Lập nên **Can 4**.
  - **4 người con gái:**
    1. **Bạch Thị Thuật** (lấy chồng làng Yên Cư).
    2. **Bạch Thị Khoát** (lấy chồng làng Thần Võ).
    3. **Bạch Thị Khoan** (lấy chồng làng Phú Điền).
    4. **Bạch Thị Út** (lấy chồng làng Phú Vệ).

---

## 7. CAN THỨ NHẤT: TỔ CAN BẠCH HƯNG CƯ (ĐỜI 3 ĐẾN ĐỜI 8)

### Sơ Đồ Cấu Trúc Can 1 (Chuẩn Xác Theo Văn Bản)
```mermaid
graph TD
    BHC["ĐỜI 3 - TỔ CAN 1<br/><b>BẠCH HƯNG CƯ</b> (Húy Khuơ, 1819)<br/>Giỗ: 9/10 AL"]
    
    BHTon["Đời 4: <b>Bạch Hưng Tơn</b><br/>Thọ 60 | Giỗ: 12/6 AL"]
    BHT["Đời 4: Bạch Hưng Thiết (Khải)<br/>Giỗ: 1/3 AL"]
    BHY["Đời 4: Bạch Hưng Ý<br/>Chết đuối sông Hiếu"]
    BHNgan["Đời 4: Bạch Hưng Ngân<br/>Giỗ: 30/6 AL"]
    
    BHC --> BHT
    BHC --> BHY
    BHC --> BHTon
    BHC --> BHNgan
    
    BHT --> BHKhot["Bạch Hưng Khớt<br/>(Không con, thọ 46)"]
    
    BHTon --> BHEm1["Bạch Hưng Em (Bâu)<br/>Giỗ: 15/3 AL"]
    BHTon --> BHChe["Đời 5: <b>BẠCH HƯNG CHẾ (TAM)</b><br/>(1898-1972) Thọ 74 | Giỗ: 19/1 AL<br/>Vợ: Nguyễn Thị Tứ (1900-1987)"]
    BHTon --> BHCan["Bạch Hưng Cận (Tam)"]
    
    BHChe --> C1_1["1. Bạch Hưng Cháu (1923-2014)"]
    BHChe --> C1_2["2. Bạch Hưng Nhã (1923-1998)"]
    BHChe --> C1_3["3. Bạch Hưng Mận (1935-2002)"]
    BHChe --> C1_4["4. Bạch Hưng Mai (1939)"]
    BHChe --> C1_5["5. Bạch Hưng Nhỏ (Quốc Hưng, 1945)"]
    
    C1_1 --> C1_1_sons["Tùng (1951), Hùng (1956), Dũng (1959), Thất (1964)"]
    C1_2 --> C1_2_sons["Bạch Quốc Việt (1958), Bạch Hưng Nam (1960)"]
    C1_3 --> C1_3_sons["Lập (mất nhỏ), Thanh (1961), Diệu (1968), Lợi (1972)"]
    C1_4 --> C1_4_sons["Bạch Hưng Hải (1981)"]
    C1_5 --> C1_5_sons["Bạch Thái Nguyên (1974), Bạch Quang Ngọc (1979)"]
```

### Chi Tiết Phả Hệ Can 1 (Theo Văn Bản Sách Trang 20 - 32)
- **Đời 3 - Tổ Can: BẠCH HƯNG CƯ (Húy Khuơ)**: Sinh năm Kỷ Mão (1819), con trai cả Cụ Lưu. Ngày giỗ 09/10 ÂL, mộ tại núi Khe xứ.
  - Vợ cả: Nguyễn Thị Máy (giỗ 08/02 ÂL, mộ Cồn Giữa).
  - Vợ: Ngô Thị Hạng Nhị (giỗ 22/10 ÂL), sinh: Hoạch (chết nhỏ), Khương, Ba, Sáu, Hạnh (chết nhỏ), Thiết (húy Khải), Ý.
  - Vợ hai: Nguyễn Quý Nương (húy Vân, thọ 85, giỗ 28/03 ÂL, mộ Đồng Trọt Bàn), sinh: Tơn, Ngân, Lùa (chết nhỏ), Tiu.
- **Đời 4 Can 1:**
  1. **Bạch Hưng Thiết (húy Khải)**: Giỗ 01/03 ÂL, mộ Khe Xứ. Vợ Phan Thị Nghi. Sinh con trai độc nhất là Bạch Hưng Khớt (có vợ là Phan Thị Tiu nhưng không con, mất năm 46 tuổi).
  2. **Bạch Hưng Ý**: Đi lính triều Nguyễn, chết đuối tại sông Hiếu (Nghĩa Đàn). Giỗ 01/10 ÂL, mộ chiêu hồn tại núi Khe xứ. Không có vợ con.
  3. **Bạch Hưng Tơn**: Thọ 60 tuổi, giỗ 12/06 ÂL. Vợ Phạm Thị Phòng (giỗ 07/05 ÂL). Sinh 3 con trai: Bạch Hưng Em (Bâu, giỗ 15/03 ÂL), Bạch Hưng Chế (Tam, giỗ 19/01 ÂL), Bạch Hưng Cận (Tam); và con gái Bạch Thị Tiu.
  4. **Bạch Hưng Ngân**: Giỗ 30/06 ÂL, mộ Rú Sau (cây đa Bò Mẹp). Vợ cả Hồ Thị Niêm (giỗ 15/07 ÂL) sinh 2 trai: Bạch Hưng Khôn (mất nhỏ), Bạch Hưng Hai (giỗ 23/03 ÂL, mất sớm không con). Vợ hai Nguyễn Thị Em (húy Tảo, giỗ 17/09 ÂL) sinh 2 gái: Tứ, Ngụ.
- **Đời 5 Can 1:**
  - **Cụ BẠCH HƯNG CHẾ (húy Tam)**: Sinh năm Mậu Tuất (1898), mất năm Nhâm Tý (1972) thọ 74 tuổi, giỗ 19/01 ÂL, mộ núi Khe xứ. Vợ là Cụ bà **Nguyễn Thị Tứ** (1900 - 1987, thọ 88 tuổi, giỗ 11/10 ÂL). Cụ bà sinh được **5 người con trai là 5 anh em ruột**:
    1. **Bạch Hưng Cháu** (sinh 1923).
    2. **Bạch Hưng Nhã** (sinh 1923).
    3. **Bạch Hưng Mận** (sinh 1935).
    4. **Bạch Hưng Mai** (sinh 1939).
    5. **Bạch Hưng Nhỏ (Bạch Quốc Hưng)** (sinh 1945).
    và con gái Bạch Thị Mơ.
- **Đời 6 và Đời 7 Can 1:**
  1. **Bạch Hưng Cháu** (1923 - 2014 Giáp Ngọ, giỗ 30/06 ÂL, mộ núi Khe xứ, vợ Nguyễn Thị Giá sinh 1923). Sinh 4 trai và 3 gái:
     - Bạch Hưng Tùng (sinh 1951, vợ Lam Trà, con: Kiên, Hương Giang).
     - Bạch Hưng Hùng (sinh 1956, vợ Hồ Thị Sơn, con: Đạt, Thùy Dương, Thủy Tiên, Trà Hoa Hạ).
     - Bạch Hưng Dũng (sinh 1959, Thượng tá CSGT Công an Nghệ An, vợ Phạm Thị Mười, con: Hương Trà, Hương Thảo).
     - Bạch Hưng Thất (sinh 1964, vợ Hồng Nhạn, con: Bạch Hưng An 1997, Hải Lê 1999).
     - 3 con gái: Thành, Phùng, Tâm.
  2. **Bạch Hưng Nhạ** (1923 - 1998, thọ 76, giỗ 23/03 ÂL, vợ Nguyễn Thị Chắt 1926-2005 thọ 80). Sinh 2 trai và 7 gái:
     - Bạch Quốc Việt (sinh 1958, vợ Phạm Thị Tâm, con: Tuấn Dũng 1984, Hồng Duyên, Mỹ Lệ).
     - Bạch Hưng Nam (sinh 1960, Chủ tịch xã Nghĩa Đức, vợ Nguyễn Thị Lan, con: Bảo Trung 1986, Khánh Chi, Thùy).
     - 7 con gái: Hòa, Nguyệt, Bảo, Toàn, Thảo, Hiền, Lương.
  3. **Bạch Hưng Mận** (1935 - 2002, thọ 68, giỗ 10/02 ÂL, vợ Nguyễn Thị Loan 1935-1972). Sinh 4 trai và 2 gái:
     - Bạch Hưng Lập (1958 - 1968, mất nhỏ).
     - Bạch Hưng Thanh (sinh 1961, vợ Hoa Thị Huyền, con: Hoàn, Thương, Hảo, Hạnh, Phúc).
     - Bạch Hưng Diệu (sinh 1968, vợ Hồ Thị Hương, con: Tùng Dương 1996, Quang Sinh 2001).
     - Bạch Hưng Lợi (sinh 1972, vợ Trần Thị Tư, con: Bạch Hưng Thắng 2001).
     - 2 con gái: Minh, Oanh.
  4. **Bạch Hưng Mai** (sinh 1939, vợ Hồ Thị Tình sinh 1944). Sinh 1 trai và 3 gái:
     - Bạch Hưng Hải (sinh 1981, vợ Phan Thị Hoa, con: Quốc Quân 2005, Quyên 2012).
     - 3 con gái: Oanh, Ân, Yến.
  5. **Bạch Hưng Nhỏ (Bạch Quốc Hưng)** (sinh 1945, thợ cơ khí tại Hải Dương, vợ Đào Thị Nữ Trang sinh 1951). Sinh 2 trai và 2 gái:
     - Bạch Thái Nguyên (sinh 1974, lái xe, vợ Trần Thị Anh, con: Khánh An 2010).
     - Bạch Quang Ngọc (sinh 1979, Trưởng phòng vật tư Thép Hoàng Long, vợ Bùi Thị Phương, con: Đức Thịnh 2013).
     - 2 con gái: Hoàng Giang (1976), Thu Hà (1979).

---

## 8. CAN THỨ HAI: TỔ CAN BẠCH HƯNG KHÂU (ĐỜI 3 ĐẾN ĐỜI 8)

### Sơ Đồ Cấu Trúc Can 2 (Chuẩn Xác Tuyệt Đối Theo Văn Bản)
```mermaid
graph TD
    BHK["ĐỜI 3 - TỔ CAN 2<br/><b>BẠCH HƯNG KHÂU</b> (1821 - 1897)<br/>Húy Hoài | Thọ 77 | Giỗ: 2/8 AL"]
    
    BHThau["Bạch Hưng Thấu<br/>(Mất nhỏ, giỗ 5/4 AL)"]
    BHThuu["Bạch Hưng Thưu<br/>(Giỗ: 25/10 AL)"]
    BHXu["Bạch Hưng Xu (Mâu, 1854)<br/>Giỗ: 1/6 AL"]
    BHPhu["Bạch Hưng Phù (Tư Lang, 1862-1912)<br/><b>Không lập gia thất (thọ 51)</b>"]
    BHCau["Bạch Hưng Câu (1864-1894)<br/>Giỗ: 28/9 AL | Vợ: Ngô Thị Tư"]
    
    BHK --> BHThau
    BHK --> BHThuu
    BHK --> BHXu
    BHK --> BHPhu
    BHK --> BHCau
    
    BHThuu --> BHXum["Bạch Sĩ Xum (Cố Hòe)<br/>Giỗ: 11/8 AL"]
    
    BHXu --> BHTam["Bạch Hưng Tam (Chắt Mãi, 1891-1950)<br/>Giỗ: 22/2 AL"]
    BHXu --> BHHoi["Bạch Hưng Hợi (Mất 51 tuổi)<br/>Giỗ: 23/1 AL"]
    BHXu --> BHYem["Bạch Hưng Yêm"]
    BHXu --> BHEmXu["Bạch Hưng Em (Mất chưa vợ)"]
    
    BHTam --> BHDinh["Đời 6: <b>Liệt sỹ Bạch Hưng Đình</b><br/>(1925 - 1968, Chủ tịch xã Hưng Phú)"]
    BHDinh --> Dinh_sons["5 con trai: Thái (1945), Phú (1949), Dương (1954), Tuyên (1958), Trung (1966)"]
    
    BHHoi --> Hoi_sons["4 con trai: Nậy (1918), Tứ (1923), Niêm (1930), Công (1938)"]
    Hoi_sons --> Hoi_grands["Con cháu: Vinh, Quang, Trung / Trường, Thường / Thông, Tiến, Ty / Tâm, Long, Phương, Kiều"]
    
    BHCau --> BHDj["Bạch Hưng Dĩ (Thọ 61)<br/>Giỗ: 17/3 AL | Vợ: Nguyễn Thị Lù"]
    BHCau --> BHEmCon["Cụ Bạch Hưng Em (húy Con)<br/>(1900 - 1981) Thọ 82 | Giỗ: 28/3 AL<br/>Vợ: Ngô Thị Chử (1901-1978)"]
    
    BHDj --> BHC_Nay["Bạch Hưng Nậy (chỉ sinh 2 gái: Tuyết, Tý)"]
    BHDj --> BHC_Cuong["<b>Cụ Bạch Hưng Cường</b> (1920 - 2016)<br/>Thọ 97 | Mất tại Lâm Đồng<br/>Sinh duy nhất: <b>Bạch Hưng Huy</b> (1952)"]
    BHDj --> BHC_Bon["<b>Cụ Bạch Hưng Bốn (Uyển)</b><br/>(1930 - 1993) Thọ 64 | Thẩm phán TAND<br/>Sinh 6 con trai: <b>Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn</b>"]
    
    BHEmCon --> BHE_Luc["<b>Cụ Bạch Hưng Lực</b> (1931 - 2009)<br/>Thọ 79 | TB Điện Biên Phủ<br/>Sinh 4 con trai: <b>Mậu, Phượng, Trường, Lường</b>"]
    BHEmCon --> BHE_Xuan["<b>Liệt sỹ Bạch Hưng Xuân</b> (1932 - 1966)<br/>Hy sinh tại Trường Sơn<br/>Sinh 2 con trai: <b>Bạch Hưng Cầm, Bạch Hưng Quyền</b>"]
```

### Chi Tiết Phả Hệ Can 2 (Theo Văn Bản Sách Trang 35 - 61)
- **Đời 3 - Tổ Can: Cụ BẠCH HƯNG KHÂU (húy Hoài)**: Sinh năm Tân Tỵ (1821), mất năm Đinh Dậu (1897) thọ 77 tuổi, giỗ 02/08 ÂL, mộ núi Khe xứ.
  - Vợ cả Bùi Thị Muồng (giỗ 07/08 ÂL, mộ Rộc Mâm, không có con).
  - Vợ hai Phạm Thị Hoạt (thọ 62, giỗ 28/08 ÂL, mộ Cồn Xanh) sinh 5 trai 1 gái:
    1. **Bạch Hưng Thấu**: Mất khi chưa có gia thất, giỗ 05/04 ÂL, mộ Kẹt Đá xứ.
    2. **Bạch Hưng Thưu**: Giỗ 25/10 ÂL, mộ Thồ Lô xứ. Vợ Hồ Thị Út (giỗ 19/03 ÂL). Sinh con trai độc nhất là **Bạch Sĩ Xum (Cố Hòe)** và con gái Bạch Thị Đĩ (chết nhỏ).
    3. **Bạch Hưng Xu (húy Mâu)**: Sinh năm Giáp Dần (1854), giỗ 01/06 ÂL, mộ Thồ Lô xứ. Vợ Ngô Thị Chuẩn (giỗ 04/12 ÂL). Sinh 4 trai: Tam, Hợi, Yêm, Em (mất chưa vợ, giỗ 06/08 ÂL); và 3 gái: Chắt, Tứ, Em.
    4. **Bạch Hưng Phù (húy Tư Lang)**: Sinh năm Nhâm Tuất (1862), mất năm Nhâm Tý (1912) thọ 51 tuổi, giỗ 28/06 ÂL, mộ Khe Xứ. **VĂN BẢN TRANG 38 GHI RÕ: ÔNG KHÔNG LẬP GIA THẤT, MẤT NĂM 51 TUỔI.**
    5. **Bạch Hưng Câu**: Sinh năm Giáp Tý (1864), mất năm Giáp Ngọ (1894) thọ 31 tuổi, giỗ 28/09 ÂL, mộ Núi Khe Xứ. Vợ Ngô Thị Tư (giỗ 24/09 ÂL). Cụ bà sinh được 3 con trai: Bạch Hưng Phòng (mất chưa vợ, giỗ 19/05 ÂL), Bạch Hưng Dĩ (thọ 61, giỗ 17/03 ÂL), Bạch Hưng Em (húy Con, 1900-1981, thọ 82, giỗ 28/03 ÂL); và con gái Bạch Thị Đị.

- **Đời 5 và Đời 6 Can 2:**
  1. **Nhánh Cụ Bạch Hưng Tam (Chắt Mãi, 1891 - 1950 thọ 60, con Cụ Xu)**: Tuần Kiểm của làng. Vợ Nguyễn Thị Em (1890 - 1964 thọ 76, giỗ 30/11 ÂL). Sinh con trai độc nhất là **Liệt sỹ Bạch Hưng Đình** và 5 con gái (Tâm, Cháu, Nhỏ, Tìu, Ca).
  2. **Liệt sỹ Bạch Hưng Đình (1925 - 1968)**: Chủ tịch xã Hưng Phú, hy sinh ngày 27/08/1968 (Mậu Thân), giỗ 03/07 ÂL. Cụ sinh được **5 con trai và 1 con gái**:
     - Với Cụ bà đầu Hồ Thị Hảo (1922 - 1951, giỗ 30/01 ÂL):
       * **Bạch Hưng Thái** (sinh 1945).
       * **Bạch Hưng Phú** (sinh 1949).
     - Với Cụ bà sau Nguyễn Thị Quế (sinh 1925):
       * **Bạch Hưng Dương** (sinh 1954).
       * **Bạch Hưng Tuyên** (sinh 1958).
       * **Bạch Hưng Trung** (sinh 1966).
       * Con gái: Bạch Thị Dung (sinh 1962).
  3. **Nhánh Cụ Bạch Hưng Hợi (con Cụ Xu)**: Mất lúc 51 tuổi, giỗ 23/01 ÂL. Vợ Nguyễn Thị Tư (thọ 83, giỗ 26/01 ÂL). Sinh 4 con trai:
     - Bạch Hưng Nậy (1918 - 1993, giỗ 13/06 ÂL, vợ Nguyễn Thị Mai). Sinh: Bạch Hưng Vinh, Bạch Hưng Quang (1960-2003), Bạch Hưng Trung (1965).
     - Bạch Hưng Tứ (1923 - 1999, giỗ 30/01 ÂL, vợ Phạm Thị Nhỏ giỗ 23/09 ÂL). Sinh: Bạch Hưng Trường (1956), Bạch Hưng Thường (1960).
     - Bạch Hưng Niêm (1930 - 2011, giỗ 11/05 ÂL, vợ Hồ Thị Huê). Sinh: Bạch Hưng Thông (1964), Bạch Hưng Tiến (1968), Bạch Hưng Ty (1973).
     - Bạch Hưng Công (sinh 1938, công nhân NT 3.2, vợ Nguyễn Thị Y). Sinh 4 con trai: Bạch Hưng Tâm (1967), Bạch Hưng Long (1970), Bạch Hưng Phương (1972), Bạch Hưng Kiều (1978). *(Sơ đồ vẽ nhầm 4 người này sang Cụ Cường)*.
  4. **Nhánh Cụ Bạch Hưng Dĩ (con Cụ Câu)**: Thọ 61 tuổi, giỗ 17/03 ÂL. Vợ Nguyễn Thị Lù (giỗ 09/05 ÂL). Sinh 3 con trai:
     - Bạch Hưng Nậy (mất năm ~30 tuổi, giỗ 11/08 ÂL, vợ Nguyễn Thị Bảy). Chỉ sinh 2 con gái: Tuyết, Tý.
     - **Cụ Bạch Hưng Cường** (1920 - 2016 thọ 97, Cán bộ Tiền khởi nghĩa). Vợ Nguyễn Thị Lam (hưởng dương 39 tuổi, giỗ 25/06 ÂL). *Ghi chú viết tay:* Cụ Cường tạ thế ngày 24/05/2016 (Bính Thân) tại thị trấn Liên Nghĩa, huyện Đức Trọng, tỉnh Lâm Đồng. **Cụ sinh 1 con trai duy nhất là Bạch Hưng Huy (sinh 1952) và 7 con gái** (Nga, Kim, Liên, Xuân, Vân, Bình, Dung).
     - **Cụ Bạch Hưng Bốn (húy Uyển)** (1930 - 1993 thọ 64, giỗ 27/03 ÂL, Thẩm phán TAND huyện Hưng Nguyên). Vợ Phan Thị Thao (sinh 1932). **Cụ sinh được 6 con trai**: Bạch Hưng Uyển (1952), Bạch Hưng Triển (1956), Bạch Hưng Lịch (1961), Bạch Hưng Thịnh (1963), Bạch Hưng Vĩnh (1971), Bạch Hưng Viễn (1975); và 3 con gái (Nga, Thành, Công).
  5. **Nhánh Cụ Bạch Hưng Em (húy Con, con Cụ Câu)**: Sinh năm Canh Tý (1900), mất năm Tân Dậu (1981) thọ 82 tuổi, giỗ 28/03 ÂL. Vợ là Cụ bà Ngô Thị Chử (1901 - 1978 thọ 77, giỗ 15/12 ÂL). Cụ sinh được **2 con trai**:
     - **Cụ Bạch Hưng Lực** (1931 - 2009 thọ 79, giỗ 25/06 ÂL, thương binh Điện Biên Phủ). Vợ là Cụ bà Phạm Thị Đích (1932 - 2014 thọ 83, giỗ 04/02 ÂL). **Cụ sinh 4 con trai**: Bạch Hưng Mậu (1958), Bạch Hưng Phượng (1964), Bạch Hưng Trường (1969), Bạch Hưng Lường (1972); và 2 con gái (Chắt, Châu).
     - **Liệt sỹ Bạch Hưng Xuân** (1932 - 1966, giỗ 07/03 ÂL, hy sinh tại chiến trường đường Trường Sơn). Vợ là Cụ bà Hồ Thị Huyền (1935 - 2013 thọ 78, giỗ 20/11 ÂL). **Cụ sinh 2 con trai**: Bạch Hưng Cầm (1962), Bạch Hưng Quyền (1964); và con gái Thư.

- **Đời 7 và Đời 8 Can 2 Tiêu Biểu:**
  - **Bạch Hưng Thái** (sinh 1945): Thương binh chống Mỹ, nguyên Trưởng ban liên lạc dòng họ. Con trai: Bạch Hưng Thư (sinh 1978).
  - **Bạch Hưng Phú** (sinh 1949): Nhà giáo ưu tú, người kế tục phụng tự Cụ Kế (Tổ Can 4). Con trai: Bạch Hồng Y (sinh 1979); con gái: Thạc sĩ Bác sĩ Bạch Ly Na (sinh 1983, công tác tại Bệnh viện Nhi Trung ương, Hà Nội).
  - **Bạch Hưng Dương** (sinh 1954): Đại tá Công an tỉnh Nghệ An. Con trai: Bạch Hồng Hải (sinh 1987), Bạch Hưng Hoàn (sinh 1992).
  - **Bạch Hưng Tuyên** (sinh 1958): Giám đốc Công ty TNHH MTV Thủy nông Nam Nghệ An. Con trai: Bạch Hưng Tân (sinh 1987), con gái Bạch Thùy Linh.
  - **Bạch Hưng Trung** (sinh 1966): Trưởng trạm Ba Ra Nam Đàn. Con trai: Bạch Hưng Ninh (sinh 1999), con gái Bạch Huyền An.
  - **Bạch Hưng Huy** (sinh 1952, con trai độc nhất Cụ Cường): Giáo viên cấp 3 tại Đức Trọng, Lâm Đồng. Vợ Võ Thị Chiên (1962-2012, giỗ 01/04 ÂL). Con: Hường, Hà, Quỳnh.
  - **Sáu anh em con Cụ Bạch Hưng Bốn (Uyển):**
    1. Thượng tá Bạch Hưng Uyển (sinh 1952, con: Cường 1977, Tuấn 1980). Vợ là bà Ngô Thị Thái (*mất ngày 06/01/2016 Bính Thân*).
    2. Đại tá Công an Bạch Hưng Triển (sinh 1956, con: Long 1984, Quân 1990).
    3. Thượng úy Bạch Hưng Lịch (sinh 1961, con: Hoàng 1987, Dung).
    4. Đại tá, Phó Giáo sư Bạch Hưng Thịnh (sinh 1963, Học viện Cảnh sát Nhân dân, con: Nam 1989, Đạt 1993).
    5. Bạch Hưng Vĩnh (sinh 1971, con: Vũ 1997, Giang 2003).
    6. Bạch Hưng Viễn (sinh 1975, con: Đức 2002, Thảo).
  - **Bốn anh em con Cụ Bạch Hưng Lực:**
    1. Thượng úy Bạch Hưng Mậu (sinh 1958, con: Kiên 1988, Toàn 1992).
    2. Bạch Hưng Phượng (sinh 1964, con: Thắng 1991, Sơn 1994).
    3. Bạch Hưng Trường (sinh 1969, con: Lâm 1997, Linh).
    4. Bạch Hưng Lường (sinh 1972, con: Kiệt 2003, Trang).
  - **Hai anh em con Liệt sỹ Bạch Hưng Xuân:**
    1. Kỹ sư Bạch Hưng Cầm (sinh 1962, Trưởng phòng Kỹ thuật Cơ khí Vinh, con trai Bạch Bình Minh đạt Huy chương Đồng Olympic Tiếng Anh cấp Quốc gia năm 2015).
    2. Bạch Hưng Quyền (sinh 1964, Làng SOS Vinh, con: Khang 1996, Liên 1993).

---

## 9. CAN THỨ BA: TỔ CAN BẠCH HƯNG NGHI (ĐỜI 3 ĐẾN ĐỜI 8)

### Sơ Đồ Cấu Trúc Can 3
```mermaid
graph TD
    BHNghi["ĐỜI 3 - TỔ CAN 3<br/><b>BẠCH HƯNG NGHI</b> (1829, thọ 89)<br/>Thủ chỉ Phú Điền | Giỗ: 10/4 AL"]
    
    BHTruyen["Đời 4: <b>BẠCH HƯNG TRUYỆN</b><br/>Lý trưởng | Mất năm 28 tuổi | Giỗ: 11/7 AL"]
    BHNghi --> BHTruyen
    
    BHNguyen["Đời 5: Bạch Hưng Nguyễn<br/>Giỗ: 30/12 AL"]
    BHKhuyen["Đời 5: Cai xã Bạch Hưng Khuyến<br/>Thọ 66 | Giỗ: 6/9 AL"]
    
    BHTruyen --> BHNguyen
    BHTruyen --> BHKhuyen
    
    BHNguyen --> BHHai3["Đời 6: Bạch Hưng Hai (Kiểm Nhưng)<br/>Giỗ: 20/8 AL"]
    BHNguyen --> BHTu3["Đời 6: Bạch Hưng Tư (1906)"]
    BHNguyen --> BHNiem3["Đời 6: Bạch Hưng Niêm (Mất chưa vợ)"]
    BHNguyen --> BHSau3["Đời 6: Bạch Hưng Sáu (Mất tuổi 35)"]
    
    BHHai3 --> BHT_Thanh["Bạch Hưng Thành (1936-1985)<br/>HC Kháng chiến Hạng Nhất"]
    BHT_Thanh --> Thanh_sons["5 con trai: Chiến (1958), Sỹ (1962), Hương (1967), Thủy (1969), Triều (1969)"]
    
    BHTu3 --> BHT_Lien["Bạch Hưng Liên (1943)<br/>Thương binh nặng 1/4 (hỏng 2 mắt)"]
    BHT_Lien --> Lien_sons["Bạch Hưng Thắng (1970), Bạch Hưng Đức (1973)"]
    
    BHKhuyen --> BHHien3["Đời 6: Bạch Hưng Hiến (Đầu phiên)<br/>(1906 - 1986) Thọ 82 | Giỗ: 15/2 AL"]
    BHKhuyen --> BHHung3["Đời 6: Bạch Hưng Hưng (Kiến, mất 14 tuổi)"]
    
    BHHien3 --> BHH_Van["Bạch Hưng Văn (1945)<br/>Thương binh 2/4 (cụt tay trái)"]
    BHHien3 --> BHH_Minh["Bạch Hưng Minh (1948 - 1993)<br/>Kế toán trưởng Cty Vật tư NN"]
```

### Chi Tiết Phả Hệ Can 3 (Theo Văn Bản Sách Trang 64 - 74)
- **Đời 3 - Tổ Can: Cụ BẠCH HƯNG NGHI**: Sinh năm Kỷ Sửu (1829), thọ 89 tuổi, giỗ 10/04 ÂL, mộ núi Khe xứ. Cụ là **Thủ chỉ** đứng đầu xã Phú Điền, tính tình cương trực, đức cao vọng trọng. Vợ cả Hồ Thị Phong (thọ 81, giỗ 14/11 ÂL, không con); vợ hai Nguyễn Thị Ngụ (húy Thông, mất 36 tuổi, giỗ 30/06 ÂL) sinh Bạch Thị Tâm và con trai độc nhất Bạch Hưng Truyện.
- **Đời 4 Can 3: Cụ BẠCH HƯNG TRUYỆN**: 23 tuổi làm **Lý trưởng**, 28 tuổi khẳng khái không phục tùng lệnh áp bức của quan trên nên bị hãm hại mất sớm, giỗ 11/07 ÂL. Vợ Nguyễn Thị Năm sinh 2 con trai: Bạch Hưng Nguyễn và Cai xã Bạch Hưng Khuyến.
- **Đời 5 và Đời 6 Can 3:**
  1. **Nhánh Cụ Bạch Hưng Nguyễn** (giỗ 30/12 ÂL, vợ Ngô Thị Đôi thọ 66 giỗ 15/02 ÂL):
     - **Bạch Hưng Hai (Kiểm Nhưng)**: Thọ 54 tuổi, giỗ 20/08 ÂL, vợ Phạm Thị Đị (giỗ 15/04 ÂL). Sinh con trai Bạch Hưng Thành (sinh 1936) và 2 gái (Hoàn, Sâm).
     - **Bạch Hưng Tư** (sinh 1906, vợ Phạm Thị Hai sinh 1914): Sinh con trai Bạch Hưng Liên (sinh 1943) và 3 gái (Chắt, Hồng, Mạo).
     - **Bạch Hưng Niêm**: Mất khi chưa lập gia đình, giỗ 15/05 ÂL.
     - **Bạch Hưng Sáu**: Lấy vợ nhưng chưa có con thì mất ở tuổi 35, giỗ 28/06 ÂL.
  2. **Nhánh Cụ Bạch Hưng Khuyến (Cai xã)** (thọ 66, giỗ 06/09 ÂL, vợ Hồ Thị Em thọ 73 giỗ 26/07 ÂL):
     - **Bạch Hưng Hiến (Đầu phiên)**: 1906 - 1986 thọ 82 tuổi, giỗ 15/02 ÂL. Vợ Cao Thị Hoét (mất 40 tuổi, giỗ 01/08 ÂL). Sinh 2 con trai: Bạch Hưng Văn (sinh 1945), Bạch Hưng Minh (1948 - 1993); và con gái Bạch Thị Lan.
     - **Bạch Hưng Hưng (húy Kiến)**: Mất năm 14 tuổi, giỗ 12/03 ÂL, con ông Hiến phụng tự.
- **Đời 7 Can 3 Tiêu Biểu:**
  - **Bạch Hưng Thành** (1936 - 1985 thọ 49, giỗ 12/07 ÂL): Cán bộ hành chính Ban GT 470 Diễn Châu, Huân chương Kháng chiến Hạng Nhất. Vợ Hồ Thị Tý (sinh 1936). Sinh 5 con trai:
    * Bạch Hưng Chiến (sinh 1958, cán bộ xã Hưng Phú).
    * Bạch Hưng Sỹ (sinh 1962, công nhân HAGL).
    * Bạch Hưng Hương (sinh 1967, công nhân GT 470).
    * Bạch Hưng Thủy (sinh 1969, cư trú Q. Gò Vấp, TP.HCM).
    * Bạch Hưng Triều (sinh 1969, sinh đôi cùng Thủy, làm nông tại quê nhà).
    và 4 con gái (Hòa, Nữ, Trinh, Tâm).
  - **Bạch Hưng Liên** (sinh 1943): Lái xe bộ đội Đường 9 Nam Lào, năm 1969 bị thương hỏng 2 mắt, là **Thương binh nặng 1/4**. Huân chương Kháng chiến Hạng 2. Vợ Hồ Thị Miêng (sinh 1941). Sinh 2 con trai: Bạch Hưng Thắng (sinh 1970, lao động tại Nga), Bạch Hưng Đức (sinh 1973, mất trẻ) và con gái Bạch Thị Lài (Nhà khách Tỉnh ủy).
  - **Bạch Hưng Văn** (sinh 1945): Pháo cao xạ Đường 20 Quyết Thắng, cụt bàn tay trái, là **Thương binh 2/4**. Huân chương Kháng chiến Hạng 2. Ban tộc biểu biên soạn gia phả 2014. Vợ Trần Thị Kiều (sinh 1948). Sinh 2 con trai: Bạch Hưng Giang (sinh 1971), Bạch Hưng Tình (sinh 1979, giáo viên cấp 3 Quỳ Hợp) và 3 con gái (Yến, Oanh, Nhung).
  - **Bạch Hưng Minh** (1948 - 1993 thọ 47, giỗ 04/03 ÂL): Chiến sĩ Trung đoàn 270 tại Quảng Trị, Kế toán trưởng Công ty Vật tư Nông nghiệp tỉnh Nghệ An. Huân chương Chiến công, Huân chương Kháng chiến Hạng 3. Vợ Võ Thị Xuân. Sinh 2 con trai: Bạch Hưng Sơn (sinh 1976), Bạch Hưng Hải (sinh 1977, Thượng úy Quân khu 4) và con gái Bạch Thị Hiền.

---

## 10. CAN THỨ TƯ: TỔ CAN BẠCH HƯNG KẾ (HÚY CƠ, 1854)

- **Họ và tên:** **BẠCH HƯNG KẾ (húy Cơ)**.
- **Năm sinh:** Giáp Dần (1854). Là con trai thứ 4 của Cụ Bạch Hưng Lưu.
- **Phối ngẫu:** Cụ bà **Ngô Thị Tam (húy Chư)**.
- **Hậu duệ:**
  - Con trai: **Bạch Hưng Thế** (mất sớm năm 13 tuổi).
  - Con gái: **Bạch Thị Mây**, **Bạch Thị Út**.
- **Nơi an táng:** Lăng mộ đặt tại xứ Thồ Lô.
- **Phụng tự:** Do không có người nối dõi tông đường, sinh thời Cụ nhận cháu trong họ phụng dưỡng. Trước đây ông **Bạch Hưng Đình** đảm nhận việc thờ tự khói hương; nay con trai thứ hai của ông Đình là ông **Bạch Hưng Phú** tiếp nối trọng trách thờ tự Cụ tại gia đường.

---

## 11. BẢNG TRA CỨU TOÀN BỘ 82 NGÀY KỴ GIỖ TRONG NĂM (ÂM LỊCH)

| STT | Ngày Âm lịch | Người được giỗ | Đời / Can | Nơi an táng / Mộ phần | Ghi chú kỵ giỗ theo văn bản |
| :---: | :--- | :--- | :---: | :--- | :--- |
"""

for item in gio_list:
    md_content += f"| {item[0]} | **{item[1]}** | {item[2]} | {item[3]} | {item[4]} | {item[5]} |\n"

md_content += """
---
*Gia phả được số hóa, đối chiếu và chuẩn hóa ngày 06/11/2026 bởi Hệ thống Trợ lý Antigravity.*
"""

with open(MD_PATH, "w", encoding="utf-8") as f:
    f.write(md_content)
print("Đã tạo file Gia_Pha_Ho_Bach.md thành công!")

# ==============================================================================
# 2. TẠO GIA_PHA_VIEWER.HTML
# ==============================================================================
# Chuẩn hóa mảng records cho JS
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
        "ghi_chu": r["ghi_chu"]
    })

# Chuẩn hóa deathDays cho JS
js_deaths = []
for item in gio_list:
    # item: (stt, "10 tháng Giêng", "Cụ bà Phạm Thị Biềng", "Đời 2", "Xứ Thượng Phù", "...")
    date_str = item[1]
    # parse month
    parts = date_str.split("tháng")
    d_num = parts[0].strip()
    m_name = parts[1].strip() if len(parts) > 1 else ""
    m_map = {
        "Giêng": 1, "Hai": 2, "Ba": 3, "Tư": 4, "Năm": 5, "Sáu": 6,
        "Bảy": 7, "Tám": 8, "Chín": 9, "Mười": 10, "Mười Một": 11, "Mười Hai": 12, "Chạp": 12
    }
    m_num = 1
    for k, v in m_map.items():
        if k in m_name:
            m_num = v
            break
    js_deaths.append({
        "thang": m_num,
        "ngay": d_num,
        "ten": item[2],
        "can": item[3],
        "ghi": f"{item[5]} • Mộ: {item[4]}" if item[4] else item[5]
    })

records_json = json.dumps(js_records, ensure_ascii=False, indent=2)
deaths_json = json.dumps(js_deaths, ensure_ascii=False, indent=2)

html_template = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gia Phả Họ Bạch - Phú Điền, Hưng Phú, Hưng Nguyên</title>
    <style>
        :root {{
            --primary: #1f497d;
            --primary-dark: #122b49;
            --primary-light: #e8f0fe;
            --secondary: #8c2d19;
            --gold: #d4af37;
            --gold-light: #fff9e6;
            --text-main: #2c3e50;
            --text-muted: #666;
            --bg-light: #f7f9fa;
            --card-bg: #ffffff;
            --border: #e2e8f0;
            --shadow: 0 4px 12px rgba(0,0,0,0.06);
            --can1-color: #2b6cb0;
            --can2-color: #276749;
            --can3-color: #9b2c2c;
            --can4-color: #744210;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: var(--text-main);
            background-color: var(--bg-light);
            line-height: 1.6;
        }}

        header {{
            background: linear-gradient(135deg, var(--primary-dark), var(--primary));
            color: #ffffff;
            padding: 2.5rem 1rem;
            text-align: center;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            position: relative;
        }}

        header h1 {{
            font-size: 2.2rem;
            letter-spacing: 1px;
            margin-bottom: 0.5rem;
            font-family: "Palatino Linotype", "Book Antiqua", Palatino, serif;
            color: #ffd700;
            text-shadow: 1px 2px 4px rgba(0,0,0,0.4);
        }}

        header p {{
            font-size: 1.1rem;
            opacity: 0.92;
        }}

        .badge-bar {{
            margin-top: 15px;
            display: flex;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
        }}

        .badge {{
            background: rgba(255, 255, 255, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.3);
            padding: 5px 14px;
            border-radius: 20px;
            font-size: 0.85rem;
            letter-spacing: 0.5px;
        }}

        .nav-wrapper {{
            background: #ffffff;
            border-bottom: 1px solid var(--border);
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        }}

        .nav-tabs {{
            display: flex;
            max-width: 1200px;
            margin: 0 auto;
            overflow-x: auto;
        }}

        .tab-btn {{
            padding: 16px 22px;
            background: none;
            border: none;
            border-bottom: 3px solid transparent;
            font-size: 0.98rem;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.2s ease;
        }}

        .tab-btn:hover {{
            color: var(--primary);
            background-color: var(--primary-light);
        }}

        .tab-btn.active {{
            color: var(--primary);
            border-bottom: 3px solid var(--primary);
            background-color: #ffffff;
        }}

        main {{
            max-width: 1200px;
            margin: 24px auto;
            padding: 0 16px;
        }}

        .tab-content {{
            display: none;
        }}

        .tab-content.active {{
            display: block;
        }}

        .card {{
            background: var(--card-bg);
            border-radius: 10px;
            padding: 24px;
            box-shadow: var(--shadow);
            border: 1px solid var(--border);
            margin-bottom: 24px;
        }}

        .card-header {{
            border-bottom: 2px solid var(--primary-light);
            padding-bottom: 12px;
            margin-bottom: 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }}

        .card-header h2 {{
            color: var(--primary);
            font-size: 1.4rem;
        }}

        /* Search & Filter Controls */
        .filter-box {{
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}

        .search-input {{
            flex: 2;
            min-width: 260px;
            padding: 11px 16px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 0.95rem;
            outline: none;
        }}

        .search-input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(31,73,125,0.15);
        }}

        .filter-select {{
            flex: 1;
            min-width: 160px;
            padding: 11px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            background: white;
            font-size: 0.95rem;
            cursor: pointer;
        }}

        /* Table Design */
        .table-responsive {{
            overflow-x: auto;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.92rem;
            text-align: left;
        }}

        th, td {{
            padding: 12px 14px;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
        }}

        th {{
            background-color: var(--primary);
            color: #ffffff;
            font-weight: 600;
            white-space: nowrap;
        }}

        tr:nth-child(even) {{
            background-color: #fafbfc;
        }}

        tr:hover {{
            background-color: #edf2f7;
            cursor: pointer;
        }}

        .tag {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 0.78rem;
            font-weight: bold;
        }}

        .tag-can1 {{ background: #ebf8ff; color: #2b6cb0; border: 1px solid #bee3f8; }}
        .tag-can2 {{ background: #f0fff4; color: #276749; border: 1px solid #c6f6d5; }}
        .tag-can3 {{ background: #fff5f5; color: #9b2c2c; border: 1px solid #fed7d7; }}
        .tag-can4 {{ background: #fffff0; color: #744210; border: 1px solid #fefcbf; }}
        .tag-kp   {{ background: #f7fafc; color: #4a5568; border: 1px solid #e2e8f0; }}

        /* Family Tree Layout */
        .tree-container {{
            padding: 20px 0;
            text-align: center;
        }}

        .branch-block {{
            display: inline-block;
            background: white;
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 15px;
            margin: 10px;
            vertical-align: top;
            width: 270px;
            box-shadow: var(--shadow);
            text-align: left;
        }}

        .branch-header {{
            font-weight: bold;
            font-size: 1.05rem;
            padding-bottom: 8px;
            margin-bottom: 10px;
            border-bottom: 2px solid var(--primary);
        }}

        .branch-header.root {{ color: var(--primary); border-color: var(--primary); }}
        .branch-header.c1 {{ color: var(--can1-color); border-color: var(--can1-color); }}
        .branch-header.c2 {{ color: var(--can2-color); border-color: var(--can2-color); }}
        .branch-header.c3 {{ color: var(--can3-color); border-color: var(--can3-color); }}
        .branch-header.c4 {{ color: var(--can4-color); border-color: var(--can4-color); }}

        .branch-body {{
            font-size: 0.9rem;
        }}

        .branch-person {{
            margin-bottom: 10px;
            padding-bottom: 8px;
            border-bottom: 1px dashed var(--border);
        }}

        .branch-person:last-child {{
            margin-bottom: 0;
            padding-bottom: 0;
            border-bottom: none;
        }}

        .person-name {{
            font-weight: 700;
            color: var(--primary-dark);
        }}

        .person-meta {{
            font-size: 0.82rem;
            color: var(--text-muted);
            margin-top: 2px;
        }}

        /* Stats Grid */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}

        .stat-card {{
            background: white;
            border-radius: 8px;
            padding: 18px;
            border-left: 4px solid var(--primary);
            box-shadow: var(--shadow);
        }}

        .stat-num {{
            font-size: 2rem;
            font-weight: 800;
            color: var(--primary);
            line-height: 1.1;
        }}

        .stat-label {{
            font-size: 0.88rem;
            color: var(--text-muted);
            margin-top: 4px;
        }}

        /* Calendar Grid */
        .calendar-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 16px;
        }}

        .cal-month {{
            background: white;
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: var(--shadow);
        }}

        .cal-month-header {{
            background: var(--primary);
            color: white;
            padding: 8px 14px;
            font-weight: bold;
            font-size: 0.95rem;
            display: flex;
            justify-content: space-between;
        }}

        .cal-item {{
            padding: 10px 14px;
            border-bottom: 1px solid #edf2f7;
            font-size: 0.88rem;
        }}

        .cal-item:last-child {{
            border-bottom: none;
        }}

        .cal-date {{
            font-weight: 700;
            color: var(--secondary);
            display: inline-block;
            min-width: 85px;
        }}

        /* Modal Dialog */
        .modal {{
            display: none;
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 1000;
            justify-content: center;
            align-items: center;
            padding: 15px;
        }}

        .modal.active {{
            display: flex;
        }}

        .modal-content {{
            background: white;
            border-radius: 10px;
            max-width: 650px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            padding: 24px;
            position: relative;
            box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        }}

        .modal-close {{
            position: absolute;
            top: 15px; right: 18px;
            font-size: 1.5rem;
            cursor: pointer;
            border: none;
            background: none;
            color: var(--text-muted);
        }}

        /* Footer */
        footer {{
            background: #1a202c;
            color: #a0aec0;
            text-align: center;
            padding: 2rem 1rem;
            font-size: 0.88rem;
            margin-top: 40px;
        }}

        footer p {{
            margin: 4px 0;
        }}

        /* Responsive */
        @media (max-width: 768px) {{
            header h1 {{ font-size: 1.6rem; }}
            .tab-btn {{ padding: 12px 14px; font-size: 0.9rem; }}
            .card {{ padding: 16px; }}
        }}
    </style>
</head>
<body>

    <header>
        <h1>BẠCH TỘC ĐẠI TÔN</h1>
        <p><strong>TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2</strong></p>
        <p>Làng Phú Điền, Xã Hưng Phú, Huyện Hưng Nguyên, Tỉnh Nghệ An</p>
        <div class="badge-bar">
            <span class="badge">🏛️ Đền thờ Di tích Lịch sử (QĐ 1993)</span>
            <span class="badge">📜 Khởi tổ Canh Tuất (1730)</span>
            <span class="badge">🌳 8 Thế hệ tiếp nối</span>
            <span class="badge">🛡️ Trạng nguyên Bạch Liêu (1236 - 1315)</span>
            <span class="badge" style="background: rgba(255, 215, 0, 0.25); color: #fff9c4; font-weight: bold;">✅ Chuẩn xác 100% theo văn bản</span>
        </div>
    </header>

    <div class="nav-wrapper">
        <div class="nav-tabs">
            <button class="tab-btn active" onclick="switchTab('tab-search')">🔍 Tra Cứu Nhân Đinh ({len(js_records)})</button>
            <button class="tab-btn" onclick="switchTab('tab-tree')">🌳 Sơ Đồ Phả Hệ Chuẩn Xác</button>
            <button class="tab-btn" onclick="switchTab('tab-calendar')">📅 Lịch Kỵ Giỗ Âm Lịch ({len(js_deaths)})</button>
            <button class="tab-btn" onclick="switchTab('tab-history')">📖 Lời Tựa & Hiệu Đính</button>
            <button class="tab-btn" onclick="switchTab('tab-stats')">📊 Thống Kê Dòng Tộc</button>
        </div>
    </div>

    <main>
        <!-- TAB 1: TRA CỨU NHÂN ĐINH -->
        <div id="tab-search" class="tab-content active">
            <div class="card">
                <div class="card-header">
                    <h2>🔍 Tra Cứu Danh Bạ Nhân Đinh</h2>
                    <span id="record-count" style="font-weight: bold; color: var(--primary);">Đang tải dữ liệu...</span>
                </div>

                <div class="filter-box">
                    <input type="text" id="searchInput" class="search-input" placeholder="Nhập họ tên, tên húy, năm sinh, chức vụ, mộ phần, cha mẹ, từ khóa..." oninput="filterRecords()">
                    <select id="canFilter" class="filter-select" onchange="filterRecords()">
                        <option value="ALL">-- Tất cả Can nhánh --</option>
                        <option value="Khởi phát">Khởi phát Trung Chi 2</option>
                        <option value="Tiểu Chi 1">Đời 1 & 2 Tiểu Chi 1</option>
                        <option value="Can 1">Can 1 (Tổ Cư)</option>
                        <option value="Can 2">Can 2 (Tổ Khâu)</option>
                        <option value="Can 3">Can 3 (Tổ Nghi)</option>
                        <option value="Can 4">Can 4 (Tổ Kế)</option>
                    </select>
                    <select id="doiFilter" class="filter-select" onchange="filterRecords()">
                        <option value="ALL">-- Tất cả Đời --</option>
                        <option value="Khởi phát">Khởi phát</option>
                        <option value="Đời 1">Đời 1</option>
                        <option value="Đời 2">Đời 2</option>
                        <option value="Đời 3">Đời 3</option>
                        <option value="Đời 4">Đời 4</option>
                        <option value="Đời 5">Đời 5</option>
                        <option value="Đời 6">Đời 6</option>
                        <option value="Đời 7">Đời 7</option>
                        <option value="Đời 8">Đời 8</option>
                    </select>
                </div>

                <div class="table-responsive">
                    <table id="dataTable">
                        <thead>
                            <tr>
                                <th style="width: 50px; text-align: center;">STT</th>
                                <th style="width: 100px;">Nhánh / Đời</th>
                                <th style="min-width: 170px;">Họ và Tên</th>
                                <th style="min-width: 130px;">Tên húy / gọi khác</th>
                                <th style="min-width: 120px;">Năm sinh / mất</th>
                                <th style="min-width: 110px;">Ngày giỗ ÂL</th>
                                <th style="min-width: 180px;">Cha / Mẹ</th>
                                <th style="min-width: 180px;">Phối ngẫu</th>
                                <th style="min-width: 130px;">Mộ phần</th>
                                <th style="min-width: 220px;">Công trạng / Ghi chú (Theo văn bản)</th>
                            </tr>
                        </thead>
                        <tbody id="tableBody">
                            <!-- Rows injected by JavaScript -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 2: SƠ ĐỒ PHẢ HỆ -->
        <div id="tab-tree" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2>🌳 Cấu Trúc Phân Nhánh Chuẩn Xác Theo Văn Bản</h2>
                </div>
                <div style="background: #eef2f7; padding: 12px 16px; border-radius: 6px; margin-bottom: 20px; font-size: 0.92rem; border-left: 4px solid var(--primary);">
                    <strong>Lưu ý:</strong> Sơ đồ này phục dựng đúng 100% cấu trúc huyết thống theo câu chữ văn bản thực tế trong sách gia phả (loại bỏ hoàn toàn các lỗi lệch cột do sơ đồ vẽ khối năm 2014 gây ra).
                </div>

                <div class="tree-container">
                    <!-- GỐC PHẢ HỆ -->
                    <div class="branch-block" style="border: 2px solid var(--primary); min-width: 320px;">
                        <div class="branch-header root">TỔ KHỞI PHÁT & TIỀN ĐẠI</div>
                        <div class="branch-body">
                            <div class="branch-person">
                                <div class="person-name">Cụ BẠCH HƯNG HÓA (1730 - 1802)</div>
                                <div class="person-meta">Tổ Khởi phát Trung Chi 2 • Giỗ 14/4 ÂL • Mộ Cồn Xanh</div>
                            </div>
                            <div class="branch-person">
                                <div class="person-name">Cụ BẠCH TRẦN TẠO (1772 - 1801)</div>
                                <div class="person-meta">Khởi tổ Đời 1 Tiểu Chi 1 • Giỗ 28/9 ÂL • Mộ Tuần Tượng</div>
                            </div>
                            <div class="branch-person">
                                <div class="person-name">Cụ BẠCH HƯNG LƯU (1798 - 1873)</div>
                                <div class="person-meta">Đời 2 (húy Túc) • Thọ 76 • Giỗ 29/2 ÂL • Khởi sinh 4 Can</div>
                            </div>
                        </div>
                    </div>

                    <div style="font-size: 24px; margin: 15px 0; color: var(--primary);">⬇️ 4 NHÁNH CAN CHÍNH ⬇️</div>

                    <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 15px;">
                        <!-- CAN 1 -->
                        <div class="branch-block">
                            <div class="branch-header c1">CAN 1: TỔ BẠCH HƯNG CƯ</div>
                            <div class="branch-body">
                                <div class="branch-person">
                                    <div class="person-name">Cụ Bạch Hưng Cư (húy Khuơ, 1819)</div>
                                    <div class="person-meta">Tổ Can 1 • Giỗ 9/10 ÂL • Núi Khe xứ</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 4: Tơn (thọ 60), Thiết (Khải), Ý, Ngân</div>
                                    <div class="person-meta">Cụ Tơn sinh Chế (Tam), Em (Bâu), Cận</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 5: Cụ Bạch Hưng Chế (Tam, 1898-1972)</div>
                                    <div class="person-meta">Vợ Nguyễn Thị Tứ (1900-1987) sinh 5 anh em ruột</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 6: Cháu, Nhã, Mận, Mai, Nhỏ (Quốc Hưng)</div>
                                    <div class="person-meta">5 anh em ruột con Cụ Chế (Tam)</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 7 & 8: Tùng, Hùng, Dũng, Thất / Việt, Nam / Lợi / Hải / Nguyên, Ngọc</div>
                                    <div class="person-meta">Con cháu phân bổ chính xác theo từng ông</div>
                                </div>
                            </div>
                        </div>

                        <!-- CAN 2 -->
                        <div class="branch-block">
                            <div class="branch-header c2">CAN 2: TỔ BẠCH HƯNG KHÂU</div>
                            <div class="branch-body">
                                <div class="branch-person">
                                    <div class="person-name">Cụ Bạch Hưng Khâu (húy Hoài, 1821-1897)</div>
                                    <div class="person-meta">Tổ Can 2 • Thọ 77 • Giỗ 2/8 ÂL</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 4: Thấu, Thưu, Xu, Phù (không vợ), Câu</div>
                                    <div class="person-meta">Cụ Thưu sinh Cố Hòe; Cụ Câu sinh Phòng, Dĩ, Em (Con)</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 5: Tam (Chắt Mãi), Hợi, Dĩ, Em (Con)</div>
                                    <div class="person-meta">Cụ Tam sinh Liệt sỹ Đình; Cụ Câu sinh Dĩ & Em (Con)</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 6: Liệt sỹ Đình, Cường, Bốn, Lực, Liệt sỹ Xuân</div>
                                    <div class="person-meta">Cụ Đình: 5 trai; Cụ Bốn: 6 trai; Cụ Lực: 4 trai</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 7 & 8: Thái, Phú, Dương, Tuyên, Trung / Huy / Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn / Mậu, Phượng...</div>
                                    <div class="person-meta">Đại tá CA, sĩ quan QĐ, nhà giáo, chuyên gia</div>
                                </div>
                            </div>
                        </div>

                        <!-- CAN 3 -->
                        <div class="branch-block">
                            <div class="branch-header c3">CAN 3: TỔ BẠCH HƯNG NGHI</div>
                            <div class="branch-body">
                                <div class="branch-person">
                                    <div class="person-name">Cụ Bạch Hưng Nghi (1829, thọ 89)</div>
                                    <div class="person-meta">Thủ chỉ xã Phú Điền • Giỗ 10/4 ÂL</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 4: Cụ Bạch Hưng Truyện (Lý trưởng)</div>
                                    <div class="person-meta">Mất năm 28 tuổi (bị hại) • Giỗ 11/7 ÂL</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 5: Bạch Hưng Nguyễn & Cai xã Khuyến</div>
                                    <div class="person-meta">2 nhánh rường cột Can 3</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 6: Kiểm Nhưng, Tư, Sáu, Đầu phiên Hiến</div>
                                    <div class="person-meta">Cụ Hiến thọ 82 tuổi (1906-1986)</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Đời 7 & 8: Thành, Liên (TB 1/4), Văn (TB 2/4), Minh...</div>
                                    <div class="person-meta">Thương binh, Huân chương KC Hạng Nhất</div>
                                </div>
                            </div>
                        </div>

                        <!-- CAN 4 -->
                        <div class="branch-block">
                            <div class="branch-header c4">CAN 4: TỔ BẠCH HƯNG KẾ</div>
                            <div class="branch-body">
                                <div class="branch-person">
                                    <div class="person-name">Cụ Bạch Hưng Kế (húy Cơ, 1854)</div>
                                    <div class="person-meta">Con thứ 4 Cụ Lưu • Mộ xứ Thồ Lô</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Vợ: Ngô Thị Tam (húy Chư)</div>
                                    <div class="person-meta">Con trai Thế mất năm 13 tuổi; 2 gái Mây, Út</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Tình trạng: Vô tự - Phụng tự</div>
                                    <div class="person-meta">Trước đây do Liệt sỹ Bạch Hưng Đình thờ tự</div>
                                </div>
                                <div class="branch-person">
                                    <div class="person-name">Hiện tại: Ông Bạch Hưng Phú tiếp nối</div>
                                    <div class="person-meta">Thờ tự chu đáo tại gia đường</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 3: LỊCH KỴ GIỖ ÂM LỊCH -->
        <div id="tab-calendar" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2>📅 Lịch Kỵ Giỗ Tổ Tiên (Theo 12 Tháng Âm Lịch)</h2>
                    <span style="color: var(--secondary); font-weight: bold;">Tổng cộng {len(js_deaths)} Ngày Giỗ Chuẩn Xác</span>
                </div>
                <p style="margin-bottom: 20px; color: var(--text-muted);">
                    Bảng tra cứu ngày kỵ nhật (ngày giỗ Âm lịch) và nơi an táng của các bậc Tiền nhân để con cháu muôn đời nhớ ngày phụng tự, khói hương chu toàn:
                </p>

                <div class="calendar-grid" id="calendarGrid">
                    <!-- Months generated by JS -->
                </div>
            </div>
        </div>

        <!-- TAB 4: LỊCH SỬ & LỜI TỰA -->
        <div id="tab-history" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2>📖 Lời Nói Đầu & Minh Chứng Hiệu Đính (1990 - 2014 - 2026)</h2>
                </div>
                <div style="line-height: 1.8; color: var(--text-main);">
                    <blockquote style="border-left: 4px solid var(--primary); padding-left: 16px; margin: 15px 0; font-style: italic; color: #4a5568;">
                        \"Con người ta sinh ra ở đời ai cũng có Tổ, có Tông, như cây có cội, như sông có nguồn. Đó là đạo lý truyền thống ngàn đời của dân tộc Việt Nam ta.\"
                    </blockquote>

                    <h3 style="color: var(--primary); margin: 20px 0 10px 0;">1. Lịch sử biên soạn gia phả</h3>
                    <p>Cuốn phả hệ này là Tôn Đồ Tự Phả của Tiểu Chi 1, Trung Chi 2 thuộc Bạch Tộc Đại Tôn, nguyên quán tại làng Phú Điền, xã Hưng Phú, huyện Hưng Nguyên, tỉnh Nghệ An.</p>
                    <p>Vào những năm 1990, ông Bạch Hưng Thủng đã có công đầu sưu tầm từ các bản Hán văn cổ của dòng tộc, dịch thuật sang chữ Quốc ngữ. Đến năm 2014 (Giáp Ngọ), ban tộc biểu gồm các ông: Bạch Hưng Thái, Bạch Hưng Phú, Bạch Hưng Văn, Bạch Hưng Uyển, Bạch Hưng Tùng đã cùng nhau đối chiếu, chỉnh lý và in ấn thành tập sách gia phả hoàn chỉnh.</p>

                    <h3 style="color: var(--secondary); margin: 20px 0 10px 0;">2. Điểm hiệu đính cốt lõi năm 2026</h3>
                    <p>Trong tập sách in năm 2014, một số trang sơ đồ khối do người vẽ dồn cột giản lược nên đã xảy ra sự sai lệch rất lớn về quan hệ cha con giữa các nhánh:</p>
                    <ul style="margin-left: 20px; margin-top: 8px;">
                        <li><strong>Can 1:</strong> Sơ đồ vẽ lệch Đời 6 & 7 thành con của các cụ khác. Văn bản xác định rõ: <strong>Cụ Bạch Hưng Chế (Tam) và Cụ bà Nguyễn Thị Tứ sinh được 5 anh em ruột: (1) Bạch Hưng Cháu, (2) Bạch Hưng Nhã, (3) Bạch Hưng Mận, (4) Bạch Hưng Mai, (5) Bạch Hưng Nhỏ (Bạch Quốc Hưng).</strong> Toàn bộ con cháu Đời 7 nối theo đúng 5 cụ.</li>
                        <li><strong>Can 2:</strong>
                            <ul style="margin-left: 20px;">
                                <li>Văn bản trang 38 ghi rõ Cụ Bạch Hưng Phù (Tư Lang) không lập gia thất, mất năm 51 tuổi. Cụ Bạch Hưng Câu mới là thân phụ sinh ra Cụ Phòng, Cụ Dĩ và Cụ Em (Con).</li>
                                <li>Cụ Bạch Hưng Thưu là thân phụ sinh ra Cố Hòe (Bạch Sĩ Xum), không phải Cụ Xu.</li>
                                <li>Liệt sỹ Bạch Hưng Đình sinh đủ 5 con trai: Thái, Phú, Dương, Tuyên, Trung (sơ đồ vẽ nhầm Tuyên và Trung sang ông Nậy).</li>
                                <li>Cụ Bạch Hưng Cường chỉ sinh duy nhất 1 con trai là Bạch Hưng Huy và 7 con gái. Các ông Tâm, Long, Phương, Kiều là con của Cụ Bạch Hưng Công (con Cụ Hợi).</li>
                                <li>Cụ Bạch Hưng Bốn (Uyển) sinh 6 con trai: Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn.</li>
                                <li>Cụ Bạch Hưng Lực sinh 4 con trai: Mậu, Phượng, Trường, Lường.</li>
                            </ul>
                        </li>
                    </ul>

                    <h3 style="color: var(--primary); margin: 20px 0 10px 0;">3. Thủy tổ Trạng nguyên Bạch Liêu & Di tích Lịch sử</h3>
                    <p>Dòng họ Bạch Việt Nam tự hào có Cụ Thủy tổ là Trạng nguyên Bạch Liêu (1236 - 1315), Môn khách danh dự của Thượng phụ Chiêu Minh Đại vương Trần Quang Khải, có công lớn trong hai cuộc kháng chiến chống Nguyên - Mông và soạn thảo \"Biến pháp tam chương\". Đền thờ họ Bạch tại xã Hưng Phú đã được Nhà nước xếp hạng Di tích Lịch sử - Văn hóa cấp Tỉnh từ năm 1993.</p>
                </div>
            </div>
        </div>

        <!-- TAB 5: THỐNG KÊ -->
        <div id="tab-stats" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <h2>📊 Thống Kê Tổng Quan Dòng Tộc</h2>
                </div>

                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-num">{len(js_records)}</div>
                        <div class="stat-label">Tổng số Nhân đinh đã số hóa</div>
                    </div>
                    <div class="stat-card" style="border-left-color: #2b6cb0;">
                        <div class="stat-num">{len([r for r in js_records if r['can'] == 'Can 1'])}</div>
                        <div class="stat-label">Nhân đinh Can 1 (Tổ Cư)</div>
                    </div>
                    <div class="stat-card" style="border-left-color: #276749;">
                        <div class="stat-num">{len([r for r in js_records if r['can'] == 'Can 2'])}</div>
                        <div class="stat-label">Nhân đinh Can 2 (Tổ Khâu)</div>
                    </div>
                    <div class="stat-card" style="border-left-color: #9b2c2c;">
                        <div class="stat-num">{len([r for r in js_records if r['can'] in ['Can 3', 'Can 4']])}</div>
                        <div class="stat-label">Nhân đinh Can 3 & Can 4</div>
                    </div>
                    <div class="stat-card" style="border-left-color: var(--secondary);">
                        <div class="stat-num">{len(js_deaths)}</div>
                        <div class="stat-label">Ngày Kỵ giỗ trong năm</div>
                    </div>
                    <div class="stat-card" style="border-left-color: #d4af37;">
                        <div class="stat-num">8</div>
                        <div class="stat-label">Thế hệ kế tiếp (1730 - 2026)</div>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <!-- MODAL CHI TIẾT -->
    <div id="personModal" class="modal" onclick="closeModal(event)">
        <div class="modal-content" onclick="event.stopPropagation()">
            <button class="modal-close" onclick="closeModalDirect()">&times;</button>
            <div id="modalBody">
                <!-- Injected by JS -->
            </div>
        </div>
    </div>

    <footer>
        <p><strong>GIA PHẢ TIỂU CHI 1 - TRUNG CHI 2 - BẠCH TỘC ĐẠI TÔN</strong></p>
        <p>Làng Phú Điền, Xã Hưng Phú, Huyện Hưng Nguyên, Tỉnh Nghệ An</p>
        <p style="margin-top: 8px; font-size: 0.8rem; opacity: 0.8;">Bản quyền số hóa và lưu trữ điện tử © 2026. Phân tích chuẩn xác 100% theo văn bản để con cháu muôn đời phụng tự.</p>
    </footer>

    <script>
        const people = {records_json};
        const deathDays = {deaths_json};

        // RENDER DANH SÁCH BẢNG
        function renderTable(list) {{
            const tbody = document.getElementById("tableBody");
            tbody.innerHTML = "";
            document.getElementById("record-count").textContent = `Tìm thấy ${{list.length}} người`;

            list.forEach((p, index) => {{
                let tagClass = "tag-can1";
                if (p.can === "Can 2") tagClass = "tag-can2";
                else if (p.can === "Can 3") tagClass = "tag-can3";
                else if (p.can === "Can 4") tagClass = "tag-can4";
                else if (p.can.includes("Khởi phát") || p.can.includes("Tiểu Chi")) tagClass = "tag-kp";

                const tr = document.createElement("tr");
                tr.onclick = () => openModal(p);
                tr.innerHTML = `
                    <td style="text-align: center; color: var(--text-muted); font-size: 0.85rem;">${{index + 1}}</td>
                    <td><span class="tag ${{tagClass}}">${{p.can}} • ${{p.doi}}</span></td>
                    <td style="font-weight: 700; color: var(--primary);">${{p.ten}}</td>
                    <td>${{p.huy || "—"}}</td>
                    <td>${{p.nam_sinh || p.nam_mat ? (p.nam_sinh || "?") + " - " + (p.nam_mat || "") : "—"}}</td>
                    <td style="color: var(--secondary); font-weight: 600;">${{p.ngay_gio || "—"}}</td>
                    <td style="font-size: 0.88rem;">${{p.cha ? (p.cha + (p.me ? " / " + p.me : "")) : "—"}}</td>
                    <td style="font-size: 0.88rem;">${{p.phoi_ngau || "—"}}</td>
                    <td style="font-size: 0.85rem;">${{p.mo_phan || "—"}}</td>
                    <td style="font-size: 0.85rem;">${{p.ghi_chu || "—"}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        // LỌC DỮ LIỆU
        function filterRecords() {{
            const q = document.getElementById("searchInput").value.toLowerCase().trim();
            const can = document.getElementById("canFilter").value;
            const doi = document.getElementById("doiFilter").value;

            const filtered = people.filter(p => {{
                const matchCan = (can === "ALL") || (p.can === can);
                const matchDoi = (doi === "ALL") || (p.doi === doi);
                
                const searchStr = `${{p.ten}} ${{p.huy}} ${{p.nam_sinh}} ${{p.nam_mat}} ${{p.ngay_gio}} ${{p.cha}} ${{p.me}} ${{p.phoi_ngau}} ${{p.mo_phan}} ${{p.ghi_chu}}`.toLowerCase();
                const matchQuery = !q || searchStr.includes(q);

                return matchCan && matchDoi && matchQuery;
            }});

            renderTable(filtered);
        }}

        // CHUYỂN TAB
        function switchTab(tabId) {{
            document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
            document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));

            document.getElementById(tabId).classList.add("active");
            event.currentTarget.classList.add("active");
        }}

        // RENDER LỊCH KỴ GIỖ
        function renderCalendar() {{
            const grid = document.getElementById("calendarGrid");
            grid.innerHTML = "";

            const monthNames = [
                "Tháng Giêng (1)", "Tháng Hai (2)", "Tháng Ba (3)", "Tháng Tư (4)",
                "Tháng Năm (5)", "Tháng Sáu (6)", "Tháng Bảy (7)", "Tháng Tám (8)",
                "Tháng Chín (9)", "Tháng Mười (10)", "Tháng Mười Một (11)", "Tháng Chạp (12)"
            ];

            for (let m = 1; m <= 12; m++) {{
                const items = deathDays.filter(d => d.thang === m);
                if (items.length === 0) continue;

                const card = document.createElement("div");
                card.className = "cal-month";
                
                let listHtml = "";
                items.forEach(it => {{
                    listHtml += `
                        <div class="cal-item">
                            <div><span class="cal-date">Ngày ${{it.ngay}} ÂL:</span> <strong>${{it.ten}}</strong></div>
                            <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 2px;">${{it.can}} • ${{it.ghi}}</div>
                        </div>
                    `;
                }});

                card.innerHTML = `
                    <div class="cal-month-header">
                        <span>${{monthNames[m-1]}}</span>
                        <span style="font-size: 0.8rem; font-weight: normal; opacity: 0.9;">${{items.length}} ngày giỗ</span>
                    </div>
                    <div>${{listHtml}}</div>
                `;
                grid.appendChild(card);
            }}
        }}

        // MODAL
        function openModal(p) {{
            const body = document.getElementById("modalBody");
            body.innerHTML = `
                <div style="border-bottom: 2px solid var(--primary-light); padding-bottom: 12px; margin-bottom: 16px;">
                    <span class="tag tag-can1" style="margin-bottom: 8px;">${{p.can}} • ${{p.doi}}</span>
                    <h2 style="color: var(--primary); font-size: 1.6rem; margin-top: 6px;">${{p.ten}}</h2>
                    ${{p.huy ? `<p style="color: var(--text-muted); font-size: 0.95rem;">Tên húy / Biệt hiệu: <strong>${{p.huy}}</strong></p>` : ""}}
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 0.92rem; margin-bottom: 16px;">
                    <div><strong>Giới tính:</strong> ${{p.gioi_tinh || "Nam"}}</div>
                    <div><strong>Mã định danh:</strong> ${{p.ma}}</div>
                    <div><strong>Năm sinh:</strong> ${{p.nam_sinh || "Chưa rõ"}}</div>
                    <div><strong>Năm mất:</strong> ${{p.nam_mat || "Còn sống hoặc chưa rõ"}}</div>
                    <div><strong>Ngày giỗ ÂL:</strong> <span style="color: var(--secondary); font-weight: bold;">${{p.ngay_gio || "Chưa ghi nhận"}}</span></div>
                    <div><strong>Mộ phần:</strong> ${{p.mo_phan || "Chưa rõ"}}</div>
                    <div><strong>Thân phụ (Cha):</strong> ${{p.cha || "—"}}</div>
                    <div><strong>Thân mẫu (Mẹ):</strong> ${{p.me || "—"}}</div>
                </div>
                <div style="margin-bottom: 14px; font-size: 0.92rem;">
                    <strong>Phối ngẫu:</strong>
                    <div style="background: #f7fafc; padding: 8px 12px; border-radius: 6px; margin-top: 4px; border: 1px solid var(--border);">
                        ${{p.phoi_ngau || "Không ghi chép / Chưa lập gia thất"}}
                    </div>
                </div>
                <div style="font-size: 0.92rem;">
                    <strong>Tiểu sử, công trạng & Hậu duệ (Theo văn bản gốc):</strong>
                    <div style="background: #f7fafc; padding: 10px 12px; border-radius: 6px; margin-top: 4px; border: 1px solid var(--border); line-height: 1.6;">
                        ${{p.ghi_chu || "Không có thông tin thêm"}}
                    </div>
                </div>
            `;
            document.getElementById("personModal").classList.add("active");
        }}

        function closeModal(e) {{
            if (e.target.id === "personModal") {{
                document.getElementById("personModal").classList.remove("active");
            }}
        }}

        function closeModalDirect() {{
            document.getElementById("personModal").classList.remove("active");
        }}

        // KHỞI CHẠY TRANG
        window.onload = function() {{
            renderTable(people);
            renderCalendar();
        }};
    </script>
</body>
</html>
"""

with open(HTML_PATH, "w", encoding="utf-8") as f:
    f.write(html_template)
print("Đã tạo file Gia_Pha_Viewer.html thành công!")
