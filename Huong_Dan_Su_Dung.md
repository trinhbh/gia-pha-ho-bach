# HƯỚNG DẪN QUẢN TRỊ & SỬ DỤNG BỘ DỮ LIỆU GIA PHẢ HỌ BẠCH
### TÔN ĐỒ TỰ PHẢ - TIỂU CHI 1 - TRUNG CHI 2 - BẠCH TỘC ĐẠI TÔN
**Địa danh:** Làng Phú Điền, Xã Hưng Phú, Huyện Hưng Nguyên, Tỉnh Nghệ An  
**Phiên bản số hóa:** Giáp Ngọ (2014) - Hiệu đính chuẩn hóa: 2026

---

## I. CẤU TRÚC BỘ TỆP GIA PHẢ

Bộ dữ liệu gia phả được thiết kế theo mô hình **"Một nguồn dữ liệu trung tâm - Đa định dạng đầu ra"** bao gồm:

| Tên tệp | Định dạng | Mục đích sử dụng |
| :--- | :--- | :--- |
| **`Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx`** | Microsoft Excel | **Nguồn dữ liệu gốc (Master Database)** để quản lý, tìm kiếm, nhập thêm các thế hệ con cháu tiếp nối. Gồm 5 sheet: *Toàn bộ nhân đinh*, *Can 1*, *Can 2*, *Can 3 & Can 4*, và *Bảng Kỵ Giỗ*. |
| **`Gia_Pha_Viewer.html`** | Trang web Offline | **Ứng dụng tra cứu gia phả:** Chỉ cần nhấp đúp để mở trên trình duyệt (Chrome, Edge, Cốc Cốc, Safari...). Có thể chép sang điện thoại để xem offline không cần mạng internet. |
| **`Gia_Pha_Ho_Bach.docx`** | Microsoft Word | **Tài liệu in ấn & lưu trữ:** Định dạng văn bản hành chính sang trọng, bìa chuẩn, lời nói đầu, lịch sử Trạng nguyên Bạch Liêu, tiểu sử chi tiết từng đời. Dùng để in thành sách hoặc xuất PDF. |
| **`Gia_Pha_Ho_Bach.md`** | Markdown text | **Tài liệu số hóa:** Dành cho lưu trữ lâu dài trên Git, GitHub hoặc phần mềm phả hệ, tích hợp sẵn các sơ đồ khối Mermaid chuẩn xác. |
| **`Gia_Pha_Ho_Bach_Ban_Chia_Se.zip`** | Tệp nén ZIP | **Gói chia sẻ hoàn chỉnh:** Chứa toàn bộ các file (Excel, Word, Markdown, Web HTML và công cụ 1-Click) đã được nén sẵn để gửi qua Zalo, Email cho con cháu dòng họ. |
| **`Cap_Nhat_Gia_Pha.bat`** | File thực thi (Batch) | **Công cụ đồng bộ 1-Click:** Nhấp đúp chuột để tự động chuyển toàn bộ dữ liệu mới từ Excel sang Word, Markdown và Web Viewer. |
| **`sync_from_excel.py`** | Script Python | Động cơ xử lý chuyển đổi dữ liệu tự động giữa các định dạng. |

---

## II. HƯỚNG DẪN BỔ SUNG CON CHÁU TIẾP THEO (ĐỜI 8, ĐỜI 9...)

Bạn không cần biết lập trình, chỉ cần thực hiện 2 bước đơn giản sau:

### Bước 1: Nhập thêm người vào Excel
1. Mở file [**`Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx`**](file:///d:/python/gia%20pha%20ho%20Bach/Danh_Sach_Nhan_Dinh_Ho_Bach.xlsx) bằng Excel.
2. Chọn Sheet **`Toàn bộ nhân đinh`** (hoặc Sheet của Can tương ứng).
3. Thêm dòng mới ở cuối bảng và điền các cột:
   - **STT:** Số thứ tự tiếp theo.
   - **Mã ĐD:** Đặt theo cấu trúc `<Can>-<Đời>-<STT>`. Ví dụ: `C2-D8-15`, `C1-D8-09`.
   - **Can:** Chọn đúng Can (`Can 1`, `Can 2`, `Can 3`, `Can 4`).
   - **Đời:** Ghi rõ thế hệ (`Đời 8`, `Đời 9`).
   - **Họ và Tên:** Họ tên đầy đủ.
   - **Tên Húy / Biệt danh:** Tên gọi ở nhà hoặc bí danh (nếu có).
   - **Giới tính:** `Nam` hoặc `Nữ`.
   - **Năm sinh:** Ví dụ: `2015 (Ất Mùi)`.
   - **Năm mất / Thọ:** Nếu còn sống thì để trống; nếu đã mất thì ghi năm mất và thọ.
   - **Ngày giỗ ÂL:** Ghi ngày giỗ âm lịch (nếu có) để hệ thống tự đưa vào Lịch Kỵ Giỗ.
   - **Cha:** Ghi họ tên người cha thuộc dòng họ (ví dụ: `Bạch Hưng Cường`, `Bạch Hưng Thái`, `Bạch Hưng Văn`...).
   - **Mẹ:** Ghi họ tên mẹ (nếu có).
   - **Phối ngẫu:** Ghi tên vợ/chồng, năm sinh, quê quán.
   - **Mộ phần:** Nơi an táng (nếu có).
   - **Công trạng / Ghi chú:** Trình độ học vấn, chức vụ, nghề nghiệp, nơi cư trú, thành tích...
   - **ID_CHA (Cột 16):** **(RẤT QUAN TRỌNG)** Nhập số thứ tự (STT) của người Cha. Đây là mã số liên kết giúp hệ thống gắn kết cây phả hệ chính xác 100%, không bao giờ bị nhầm lẫn giữa các vị trùng tên trong dòng họ.
4. **Lưu (Save) và ĐÓNG HẲN file Excel lại.**

### Bước 2: Chạy đồng bộ tự động 1-Click
- Nhấp đúp chuột vào file [**`Cap_Nhat_Gia_Pha.bat`**](file:///d:/python/gia%20pha%20ho%20Bach/Cap_Nhat_Gia_Pha.bat).
- Hệ thống sẽ tự động đọc từ Excel và tái tạo lại toàn bộ file **Word**, **Markdown** và **Web Viewer** trong khoảng 2–3 giây.
- Khi màn hình hiện `[THÀNH CÔNG]`, toàn bộ các file đã đồng bộ dữ liệu mới nhất.

---

## III. HƯỚNG DẪN TRA CỨU TRÊN WEB VIEWER (`Gia_Pha_Viewer.html`)

- **Cách mở:** Nhấp đúp vào file [**`Gia_Pha_Viewer.html`**](file:///d:/python/gia%20pha%20ho%20Bach/Gia_Pha_Viewer.html), trang web sẽ tự động mở trên trình duyệt (không cần kết nối internet).
- **Tính năng nổi bật:**
  1. **🔍 Tra cứu nhân đinh:** Gõ bất kỳ từ khóa nào (tên, năm sinh, tên húy, quê quán, tên cha mẹ, chức vụ...) vào ô tìm kiếm để tìm tức thì.
  2. **Bộ lọc Can & Đời:** Xem riêng từng nhánh Can (Can 1, 2, 3, 4) hoặc từng thế hệ (từ Khởi phát đến Đời 8).
  3. **Xem tiểu sử chi tiết (Hồ sơ cá nhân):** Nhấp chuột vào bất kỳ dòng nào trong bảng, cửa sổ chi tiết sẽ hiện lên đầy đủ thông tin cha mẹ, phối ngẫu, mộ phần và tiểu sử.
  4. **🌳 Sơ đồ phả hệ trực quan (Cây thế hệ):**
     - **Trên Điện thoại:** Tự động kích hoạt chế độ **📱 Cây Dọc** dạng danh mục phân cấp tối ưu, có nút `+` / `–` đóng mở từng đời gọn gàng, hỗ trợ tìm kiếm và cuộn mượt mà.
     - **Trên Máy tính:** Chế độ **🖥️ Sơ Đồ Đồ Họa** hiển thị dạng cây ngang truyền thống, cho phép cuộn chuột để phóng to/thu nhỏ, giữ chuột kéo thả (pan/zoom) và xem toàn màn hình.
     - Cho phép lọc xem nhanh theo từng Can: Toàn Tộc, Can 1, Can 2, Can 3, Can 4.
  5. **📅 Lịch Kỵ giỗ Âm lịch:** Tự động tổng hợp và sắp xếp tuần tự theo 12 tháng Âm lịch từ tháng Giêng đến tháng Chạp, hỗ trợ bộ lọc nhanh từng tháng cho cả cụ ông và cụ bà.
  6. **📖 Lời tựa & Lịch sử:** Đọc lại các di huấn cổ, lịch sử dòng họ và minh chứng hiệu đính.

---

## IV. LƯU Ý LỊCH SỬ QUAN TRỌNG VỀ ĐỘ CHÍNH XÁC

Trong bản in gia phả gốc năm 2014, một số trang sơ đồ vẽ hình khối giản lược đã xảy ra hiện tượng **lệch cột nối nhầm quan hệ cha con** giữa các nhánh. Toàn bộ cơ sở dữ liệu này đã được **đối chiếu và hiệu đính 100% dựa vào câu chữ văn bản miêu tả đích danh ở từng trang sách**:

1. **Can 1:** Cụ Bạch Hưng Chế (Tam) và Cụ bà Nguyễn Thị Tứ sinh ra **5 anh em ruột**: (1) Bạch Hưng Cháu, (2) Bạch Hưng Nhạ, (3) Bạch Hưng Mận, (4) Bạch Hưng Mai, (5) Bạch Hưng Nhỏ (Bạch Quốc Hưng). Toàn bộ con cháu Đời 7 nối theo đúng 5 cụ.
2. **Can 2:**
   - **Phân biệt 3 vị Cụ cùng tên Bạch Hưng Nậy:**
     + *Cụ Bạch Hưng Nậy (Đời 4, con Cụ Bạch Hưng Khâu):* Sinh cụ Bạch Hưng Quế.
     + *Cụ Bạch Hưng Nậy (Đời 5, con Cụ Bạch Hưng Lưỡng):* Sinh 6 người con trai gồm: Đấu, Thìn, Cẩn, Phố, Cường (Lâm Đồng), Mạc.
     + *Cụ Bạch Hưng Nậy (Đời 5, con Cụ Bạch Hưng Duệ):* Sinh 6 người con trai gồm: Mính, Diến, Lực, Bốn (Uyển), Thi, Đệ.
   - Cụ Bạch Hưng Phù (Tư Lang, 1862–1912) **không lập gia thất, mất năm 51 tuổi**. Cụ Bạch Hưng Câu mới là thân phụ sinh 3 cụ: Phòng, Dĩ, Em (Con).
   - Cụ Bạch Sĩ Xum (Cố Hòe) là con Cụ Bạch Hưng Thưu.
   - Liệt sỹ Bạch Hưng Đình (Chủ tịch xã Hưng Phú) sinh đủ **5 con trai**: Thái, Phú, Dương, Tuyên, Trung.
   - Cụ Bạch Hưng Cường (thọ 97 tuổi tại Lâm Đồng) chỉ sinh duy nhất **1 con trai là Bạch Hưng Huy** và 7 con gái. Các ông Tâm, Long, Phương, Kiều là con Cụ Bạch Hưng Công (con Cụ Hợi).
   - Cụ Bạch Hưng Bốn (Uyển) sinh 6 con trai: Uyển, Triển, Lịch, Thịnh, Vĩnh, Viễn.
   - Cụ Bạch Hưng Lực sinh 4 con trai: Mậu, Phượng, Trường, Lường.
   - Liệt sỹ Bạch Hưng Xuân sinh 2 con trai: Cầm, Quyền.
3. **Can 4:** Cụ Bạch Hưng Kế (húy Cơ, sinh 1854) vô tự, được chi họ phụng tự: trước đây là Liệt sỹ Bạch Hưng Đình, nay là ông Bạch Hưng Phú tiếp nối thờ tự chu toàn tại gia đường.
4. **Trích xuất nguyên trạng 100% văn bản sách ảnh (Cột O):** Toàn bộ 125 nhân đinh trên file Excel và các bản xuất ra đều giữ nguyên trạng từng câu chữ, bài khấn chữ Hán - Việt cổ, chức sắc và di huấn của tiền nhân từ bản dịch gia phả của ông Bạch Hưng Thủng.
5. **Chuẩn hóa cây phả hệ (Cha $\rightarrow$ Con $\rightarrow$ Cháu):** Bảng nhân đinh được sắp xếp theo đúng thứ tự duyệt cây gia hệ, ID_CHA luôn nhỏ hơn STT của con cháu, đảm bảo liên kết tuyệt đối chính xác và tương thích hoàn toàn khi chạy cập nhật tự động.

---

## V. HƯỚNG DẪN SỬ DỤNG WEB QUẢN TRỊ TRỰC TUYẾN (`https://gia-pha-ho-bach.onrender.com/`)

Hệ thống đã được nâng cấp lên Web Python (FastAPI) chạy trực tuyến, đồng bộ dữ liệu trực tiếp với **Google Sheets** và xác thực bằng **Google Sign-In**:

1. **Truy cập & Tra cứu công khai:**
   - Mở địa chỉ: **https://gia-pha-ho-bach.onrender.com/** (chờ ~20–40 giây nếu server gói miễn phí Render vừa thức dậy).
   - Mọi người đều có thể xem danh bạ, tìm kiếm theo tên/SĐT/địa chỉ/ngày giỗ và bấm **🌳 Xem Cây Phả Hệ** để xem sơ đồ cây + lịch kỵ giỗ 12 tháng Âm lịch.
2. **Đăng nhập Google để mở khóa Thêm / Sửa / Xóa:**
   - Bấm nút **Đăng nhập bằng Google** ở góc trên bên phải.
   - Tài khoản Gmail đăng nhập bắt buộc phải nằm trong sheet **`nguoi_dung`** (trạng thái `Hoạt động`). Nếu chưa có trong danh sách, hệ thống chỉ cho phép xem.
3. **Các trường thông tin mới được bổ sung:**
   - **Tách biệt Ngày mất & Ngày giỗ:** `Năm mất / Thọ`, `Ngày mất cụ thể (DL/ÂL)`, và `Ngày giỗ Âm lịch` được tách riêng biệt rõ ràng.
   - **Thông tin liên hệ hiện đại:** `Số điện thoại`, `Link Facebook`, `Nơi ở hiện tại`, `Nghề nghiệp / Học vấn`.
   - **Chọn Thân phụ theo cây:** Khi Thêm/Sửa nhân đinh, ô chọn Cha sắp xếp theo cây phả hệ và hỗ trợ tìm kiếm tiếng Việt không dấu để tự động gắn đúng `ID_CHA`.
4. **Xuất bản Web Tĩnh (`index.html`):**
   - Tại tab **🌐 Xuất Bản Lên Internet**, có tùy chọn **Ẩn thông tin riêng tư (SĐT, Link Facebook, Nơi ở)** trước khi tải về file `index.html` để đưa lên GitHub Pages / Netlify.

