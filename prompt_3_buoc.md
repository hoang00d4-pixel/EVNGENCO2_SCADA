# Prompt chuẩn - Quy trình 3 bước tổng hợp báo cáo phân xưởng

## Mục tiêu

Thực hiện đầy đủ quy trình tổng hợp báo cáo từ các file Excel trong thư mục `Data/`, gồm:

1. Đọc toàn bộ file dữ liệu `.xlsx`.
2. Tổng hợp nội dung công việc theo ngày từ tất cả phân xưởng.
3. Tạo bảng kết quả cuối cùng theo cấu trúc Excel tổng hợp.

## Prompt sử dụng

```text
Vai trò:
Bạn là cán bộ hành chính phụ trách tổng hợp báo cáo công việc cho lãnh đạo.

Bối cảnh:
Trong thư mục Data/ có nhiều file Excel định dạng .xlsx. Mỗi file tương ứng với một phân xưởng. Tên file cho biết tên phân xưởng, ví dụ PX_1.xlsx là Phân xưởng 1, PX_2.xlsx là Phân xưởng 2.

Mỗi file dữ liệu đầu vào dự kiến có 5 cột:
- STT: Số thứ tự trong file nguồn
- Ngày: Ngày làm việc
- Thời gian/Giờ: Thời gian bắt đầu ca
- Trưởng ca: Người phụ trách phân xưởng
- Nội dung: Nội dung công việc đạt được

Nhiệm vụ:

Bước 1 - Đọc dữ liệu:
1. Đọc toàn bộ file .xlsx trong thư mục Data/.
2. Xác định tên phân xưởng từ tên file.
3. Đọc sheet dữ liệu chính trong từng file.
4. Chuẩn hóa tên cột, trong đó cột "Giờ" được hiểu là "Thời gian".
5. Ghi nhận các dòng thiếu dữ liệu ở các trường quan trọng: Ngày, Thời gian/Giờ, Trưởng ca, Nội dung.

Bước 2 - Tổng hợp theo ngày:
1. Gộp dữ liệu từ tất cả phân xưởng vào một bảng chung.
2. Tổng hợp theo trường Ngày.
3. Trong từng ngày, sắp xếp theo Thời gian, sau đó theo Tên phân xưởng.
4. Không loại bỏ các dòng cùng ngày, vì mỗi phân xưởng là một đơn vị báo cáo riêng.
5. Tạo nội dung báo cáo bằng cách ghép tên trưởng ca và nội dung công việc đạt được.

Bước 3 - Tạo bảng Excel tổng hợp:
1. Tạo bảng kết quả cuối cùng với các cột:
   - STT
   - Ngày
   - Thời gian
   - Tên phân xưởng
   - Nội dung
2. Đánh lại STT từ 1 đến hết theo thứ tự sau tổng hợp.
3. Cột Nội dung phải thể hiện tên trưởng ca và nội dung công việc đạt được.
4. Bảng kết quả phải có thể ghi trực tiếp ra file Excel tổng hợp.

Ràng buộc:
- Chỉ xử lý file .xlsx trong thư mục Data/.
- Không sửa nội dung gốc của các file đầu vào.
- Không tự ý bỏ dòng dữ liệu nếu chưa ghi nhận lý do.
- Nếu thiếu Trưởng ca hoặc Nội dung, vẫn giữ dòng và ghi chú rõ:
  - Thiếu Trưởng ca: "[Thiếu trưởng ca]"
  - Thiếu Nội dung: "[Thiếu nội dung]"
- Định dạng ngày thống nhất dạng YYYY-MM-DD.
- Định dạng thời gian thống nhất dạng HH:MM:SS.
- Tên phân xưởng lấy từ tên file, ví dụ PX_1.xlsx -> Phân xưởng 1.

Quy tắc tạo cột Nội dung:
- Nếu có đủ Trưởng ca và Nội dung: "<Trưởng ca>: <Nội dung>"
- Nếu thiếu Trưởng ca: "[Thiếu trưởng ca]: <Nội dung>"
- Nếu thiếu Nội dung: "<Trưởng ca>: [Thiếu nội dung]"
- Nếu thiếu cả hai: "[Thiếu trưởng ca]: [Thiếu nội dung]"

Định dạng đầu ra:

1. Tóm tắt bước 1:
   - Danh sách file đã đọc
   - Số dòng dữ liệu của từng file
   - Cấu trúc cột phát hiện
   - Cảnh báo dữ liệu thiếu

2. Tóm tắt bước 2:
   - Số ngày làm việc
   - Số phân xưởng
   - Số dòng sau tổng hợp
   - Quy tắc sắp xếp đã áp dụng

3. Bảng Excel tổng hợp cuối cùng:
   - STT
   - Ngày
   - Thời gian
   - Tên phân xưởng
   - Nội dung
```

## Chạy thử 1 ngày

Ngày chạy thử: `2026-01-07`

### Bước 1 - Đọc dữ liệu

| File      | Tên phân xưởng | Sheet  | Dòng Excel | Ngày      | Thời gian | Trưởng ca    | Nội dung   |
| --------- | ------------------ | ------ | ----------: | ---------- | ---------- | -------------- | ----------- |
| PX_1.xlsx | Phân xưởng 1    | Sheet1 |           2 | 2026-01-07 | 08:00:00   | Nguyễn Văn A | Báo cáo 1 |
| PX_2.xlsx | Phân xưởng 2    | Sheet1 |           2 | 2026-01-07 | 08:00:00   | Đoàn Văn A  | Báo cáo 1 |
| PX_3.xlsx | Phân xưởng 3    | Sheet1 |           2 | 2026-01-07 | 08:00:00   | Trần Văn A   | Báo cáo 1 |

Kết quả bước 1: đọc được 3 dòng dữ liệu của ngày `2026-01-07` từ 3 file phân xưởng, không phát hiện thiếu dữ liệu.

### Bước 2 - Tổng hợp theo ngày

| Ngày      | Thời gian | Tên phân xưởng | Trưởng ca    | Nội dung công việc |
| ---------- | ---------- | ------------------ | -------------- | --------------------- |
| 2026-01-07 | 08:00:00   | Phân xưởng 1    | Nguyễn Văn A | Báo cáo 1           |
| 2026-01-07 | 08:00:00   | Phân xưởng 2    | Đoàn Văn A  | Báo cáo 1           |
| 2026-01-07 | 08:00:00   | Phân xưởng 3    | Trần Văn A   | Báo cáo 1           |

Kết quả bước 2: tổng hợp được 3 dòng theo cùng ngày `2026-01-07`, sắp xếp theo thời gian và tên phân xưởng.

### Bước 3 - Bảng Excel tổng hợp cuối cùng

| STT | Ngày      | Thời gian | Tên phân xưởng | Nội dung                   |
| --: | ---------- | ---------- | ------------------ | --------------------------- |
|   1 | 2026-01-07 | 08:00:00   | Phân xưởng 1    | Nguyễn Văn A: Báo cáo 1 |
|   2 | 2026-01-07 | 08:00:00   | Phân xưởng 2    | Đoàn Văn A: Báo cáo 1  |
|   3 | 2026-01-07 | 08:00:00   | Phân xưởng 3    | Trần Văn A: Báo cáo 1   |

Kết quả bước 3: bảng cuối cùng có đúng 5 cột theo yêu cầu: `STT`, `Ngày`, `Thời gian`, `Tên phân xưởng`, `Nội dung`.
Đầu ra mong muốn:

Tên file tổng hợp:
Baocao_TH_Ngaygio.xlsx
với ngaygio là ngày giờ tạo ra báo cáo

Toàn bộ nội dung font chữ Times New Roman, sử dụng cỡ chữ 12
