# Bài 3: Chuẩn hóa quy trình giao vận và tra cứu giá cước – FastMove

## Phần 1. Phân tích: Mapping 4 tham số VLOOKUP vào mô hình "Nhân viên kho tra hồ sơ"

Công ty FastMove gặp vấn đề do nhân viên phải tra cứu giá cước thủ công, gây chậm trễ khoảng 16 tiếng mỗi ngày. Để giải quyết vấn đề, công ty sử dụng hàm VLOOKUP trong Excel để tự động tìm kiếm giá cước.

Hàm VLOOKUP có cú pháp:

`=VLOOKUP(Lookup_value, Table_array, Col_index_num, Range_lookup)`

Áp dụng vào mô hình nhân viên kho tra hồ sơ:

- **Lookup_value (Giá trị cần tìm):** Mã dịch vụ trên đơn hàng mà nhân viên cần tra cứu.
- **Table_array (Vùng bảng):** Bảng danh mục giá cước chứa mã dịch vụ và giá tương ứng, giống như tủ hồ sơ lưu trữ thông tin.
- **Col_index_num (Số thứ tự cột):** Vị trí cột chứa giá cước cần lấy, giống như vị trí thông tin cần đọc trong hồ sơ.
- **Range_lookup (Kiểu tìm kiếm):** Sử dụng `FALSE` để tìm chính xác mã dịch vụ, tránh lấy nhầm giá cước.

Nhờ đó, nhân viên kho có thể tra cứu giá cước nhanh chóng, giảm thao tác thủ công và hạn chế sai sót.

## Phần 2. Triển khai: Giải pháp công nghệ cho FastMove

Giải pháp sử dụng ba công cụ Word, Excel và PowerPoint để chuẩn hóa quy trình giao vận.

### 2.1. Word – Quy trình vận hành kho

**File:** `QuyTrinhVanHanhKho.docx`

Tài liệu hướng dẫn nhân viên thực hiện quy trình vận hành kho gồm:

- Tiếp nhận đơn hàng và kiểm tra mã vận đơn.
- Phân loại hàng hóa theo khu vực giao.
- Tra cứu giá cước và kiểm tra thông tin đơn hàng.
- Bàn giao hàng hóa cho shipper và cập nhật trạng thái.

Áp dụng Styles trong Word để định dạng tiêu đề chính, tiêu đề các phần và nội dung thống nhất. Sử dụng màu xanh dương làm màu chủ đạo, màu cam để nhấn mạnh lưu ý quan trọng.

### 2.2. Excel – Tra cứu giá cước tự động

**File:** `TraCuuGiaCuoc.xlsx`

File gồm hai Sheet:

- `VanDon`: Chứa mã vận đơn, khu vực, mã dịch vụ và giá cước.
- `GiaCuoc`: Chứa mã dịch vụ, tên dịch vụ và giá cước tương ứng.

Tại ô D2 của Sheet `VanDon`, nhập công thức:

`=IFERROR(VLOOKUP(C2,GiaCuoc!$A$2:$C$4,3,FALSE),"Không tìm thấy mã dịch vụ")`

Trong đó:

- `C2`: Mã dịch vụ cần tra cứu.
- `GiaCuoc!$A$2:$C$4`: Vùng bảng giá cước ở Sheet khác.
- `$A$2:$C$4`: Cố định vùng tra cứu khi kéo công thức xuống 500 dòng.
- `3`: Lấy giá trị ở cột thứ ba của vùng tra cứu.
- `FALSE`: Tìm kiếm chính xác.
- `IFERROR`: Hiển thị thông báo thay cho lỗi công thức.

Dữ liệu Excel được chuẩn hóa theo dạng bảng, mỗi cột chứa một loại thông tin, mỗi dòng là một bản ghi, không gộp ô và không để hàng trống xen giữa dữ liệu.

### 2.3. PowerPoint – Tài liệu hướng dẫn shipper

**File:** `HuongDanShipper.pptx`

Bài trình bày gồm hai slide:

**Slide 1: Tên app**

**Slide 2: Kiểm tra thông tin giúp shipper nhận đúng đơn hàng**
- Kiểm tra mã vận đơn.
- Đối chiếu thông tin người nhận.
- Xác nhận đơn thuộc tuyến giao được phân công.

Áp dụng Miller's Law bằng cách giới hạn mỗi slide ở 3–4 nhóm ý chính, sử dụng câu ngắn gọn, hình ảnh minh họa và tiêu đề dạng kết luận. Tiêu đề sử dụng cỡ chữ tối thiểu 30 pt, nội dung khoảng 18–20 pt và màu sắc thống nhất với tài liệu Word.

## Phần 3. Kiểm thử: Xử lý lỗi #N/A khi mã vận đơn có ký tự lạ bằng IFERROR

### 3.1. Tình huống kiểm thử

Khi mã dịch vụ không tồn tại trong bảng giá hoặc chứa ký tự lạ, hàm VLOOKUP có thể trả về lỗi `#N/A`. Điều này khiến bảng tính khó đọc và gây khó khăn cho nhân viên kho.

### 3.2. Công thức xử lý lỗi

Sử dụng công thức:

`=IFERROR(VLOOKUP(C2,GiaCuoc!$A$2:$C$4,3,FALSE),"Không tìm thấy mã dịch vụ")`

IFERROR kiểm tra kết quả của công thức VLOOKUP. Nếu công thức gặp lỗi, Excel hiển thị thông báo "Không tìm thấy mã dịch vụ" thay vì hiển thị lỗi `#N/A`.

### 3.3. Bảng kiểm thử

| Dữ liệu đầu vào | Kết quả mong đợi |
|---|---|
| DV01 | 20000 |
| DV02 | 30000 |
| DV03 | 45000 |
| DV99 | Không tìm thấy mã dịch vụ |
| DV@01 | Không tìm thấy mã dịch vụ nếu mã không tồn tại trong bảng |

Các giá cước trên là dữ liệu minh họa phục vụ kiểm thử.

### 3.4. Kết quả kiểm thử

- Mã dịch vụ hợp lệ trả về đúng giá cước.
- Mã không tồn tại hiển thị thông báo thay cho lỗi `#N/A`.
- Vùng bảng tra cứu được cố định bằng dấu `$` khi kéo công thức xuống 500 dòng.
- Dữ liệu được chuẩn hóa, không gộp ô và không trùng mã dịch vụ.
- Trường hợp có khoảng trắng thừa có thể dùng TRIM để làm sạch dữ liệu trước khi tra cứu.

## Kết luận

Giải pháp giúp FastMove chuẩn hóa quy trình vận hành kho, tự động hóa việc tra cứu giá cước và hướng dẫn shipper nhận đơn thống nhất. Việc kết hợp Word, Excel và PowerPoint giúp giảm thao tác thủ công, hạn chế sai sót và hỗ trợ nâng cao hiệu suất giao vận.
