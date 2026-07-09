# Prompt chuẩn - Bước 1: Đọc toàn bộ file dữ liệu XLSX trong Data/

## Mục tiêu bước 1

Đọc toàn bộ các file dữ liệu Excel định dạng `.xlsx` trong thư mục `Data/`, xác định cấu trúc dữ liệu hiện có, kiểm tra chất lượng dữ liệu đầu vào và chuẩn bị dữ liệu cho bước tổng hợp theo ngày.

## Prompt sử dụng

```text
Vai trò:
Bạn là cán bộ hành chính phụ trách tổng hợp báo cáo công việc cho lãnh đạo.

Bối cảnh:
Trong thư mục Data/ có nhiều file dữ liệu Excel định dạng .xlsx, mỗi file tương ứng với một phân xưởng. Tên file cho biết phân xưởng, ví dụ PX_1.xlsx là Phân xưởng 1, PX_2.xlsx là Phân xưởng 2.

Mỗi file dữ liệu dự kiến có 5 cột:
- STT: Số thứ tự
- Ngày: Ngày làm việc
- Thời gian/Giờ: Thời gian bắt đầu ca
- Trưởng ca: Tên người phụ trách phân xưởng
- Nội dung: Nội dung công việc đạt được

Nhiệm vụ:
1. Đọc toàn bộ file .xlsx trong thư mục Data/.
2. Với từng file, xác định tên phân xưởng từ tên file.
3. Kiểm tra sheet dữ liệu chính và đọc toàn bộ các dòng dữ liệu.
4. Chuẩn hóa tên cột nếu cần, trong đó cột "Giờ" được hiểu là "Thời gian".
5. Kiểm tra dữ liệu thiếu ở các trường quan trọng: Ngày, Thời gian/Giờ, Trưởng ca, Nội dung.
6. Trả về bảng dữ liệu đã đọc theo cấu trúc thống nhất để phục vụ bước tổng hợp tiếp theo.

Ràng buộc:
- Chỉ xử lý file có phần mở rộng .xlsx trong thư mục Data/.
- Không tự ý bỏ qua dòng dữ liệu nếu chưa ghi nhận lý do.
- Không thay đổi nội dung gốc của các file Excel đầu vào.
- Nếu phát hiện dữ liệu thiếu, phải nêu rõ file, dòng, cột bị thiếu.
- Tên phân xưởng phải được suy ra từ tên file, ví dụ PX_1.xlsx -> Phân xưởng 1.
- Đầu ra cần rõ ràng, dễ kiểm tra, có thể dùng làm dữ liệu đầu vào cho bước 2.

Định dạng đầu ra:
1. Danh sách file đã đọc:
   - Tên file
   - Tên phân xưởng
   - Tên sheet
   - Số dòng dữ liệu
   - Số cột

2. Cấu trúc cột phát hiện:
   - Tên cột gốc
   - Tên cột chuẩn hóa
   - Ý nghĩa

3. Bảng dữ liệu chuẩn hóa:
   - STT
   - Ngày
   - Thời gian
   - Tên phân xưởng
   - Trưởng ca
   - Nội dung

4. Cảnh báo dữ liệu:
   - File
   - Dòng Excel
   - Cột
   - Vấn đề phát hiện
```

## Thử nghiệm với dữ liệu hiện có

### Danh sách file đã đọc

| Tên file | Tên phân xưởng | Sheet | Số dòng dữ liệu | Số cột |
|---|---|---:|---:|---:|
| PX_1.xlsx | Phân xưởng 1 | Sheet1 | 8 | 5 |
| PX_2.xlsx | Phân xưởng 2 | Sheet1 | 8 | 5 |
| PX_3.xlsx | Phân xưởng 3 | Sheet1 | 8 | 5 |

### Cấu trúc cột phát hiện

| Tên cột gốc | Tên cột chuẩn hóa | Ý nghĩa |
|---|---|---|
| STT | STT | Số thứ tự |
| Ngày | Ngày | Ngày làm việc |
| Giờ | Thời gian | Thời gian bắt đầu ca |
| Trưởng ca | Trưởng ca | Người phụ trách phân xưởng |
| Nội dung | Nội dung | Nội dung công việc đạt được |

### Dữ liệu mẫu sau chuẩn hóa

| STT | Ngày | Thời gian | Tên phân xưởng | Trưởng ca | Nội dung |
|---:|---|---|---|---|---|
| 1 | 2026-01-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn A | Báo cáo 1 |
| 2 | 2026-02-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn B | Báo cáo 2 |
| 1 | 2026-01-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn A | Báo cáo 1 |
| 2 | 2026-02-07 | 08:00:00 | Phân xưởng 2 |  | Báo cáo 2 |
| 1 | 2026-01-07 | 08:00:00 | Phân xưởng 3 | Trần Văn A | Báo cáo 1 |
| 5 | 2026-05-07 | 08:00:00 | Phân xưởng 3 |  |  |

### Cảnh báo dữ liệu

| File | Dòng Excel | Cột | Vấn đề phát hiện |
|---|---:|---|---|
| PX_2.xlsx | 3 | Trưởng ca | Thiếu tên trưởng ca |
| PX_3.xlsx | 6 | Trưởng ca | Thiếu tên trưởng ca |
| PX_3.xlsx | 6 | Nội dung | Thiếu nội dung công việc |

