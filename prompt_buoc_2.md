# Prompt chuẩn - Bước 2: Tổng hợp theo ngày nội dung công việc từ tất cả file dữ liệu

## Mục tiêu bước 2

Từ dữ liệu đã đọc và chuẩn hóa ở bước 1, tổng hợp nội dung công việc theo từng ngày, giữ chi tiết thời gian, tên phân xưởng, trưởng ca và nội dung công việc đạt được của từng phân xưởng.

## Prompt sử dụng

```text
Vai trò:
Bạn là cán bộ hành chính phụ trách tổng hợp báo cáo công việc cho lãnh đạo.

Bối cảnh:
Dữ liệu đầu vào là kết quả đã chuẩn hóa từ bước 1, được đọc từ toàn bộ file .xlsx trong thư mục Data/. Mỗi file tương ứng với một phân xưởng. Tên file cho biết tên phân xưởng, ví dụ PX_1.xlsx là Phân xưởng 1.

Dữ liệu đầu vào đã có các trường:
- STT
- Ngày
- Thời gian
- Tên phân xưởng
- Trưởng ca
- Nội dung

Trong đó:
- Trưởng ca là người phụ trách phân xưởng.
- Nội dung là nội dung công việc đạt được trong ca.

Nhiệm vụ:
1. Tổng hợp toàn bộ dòng dữ liệu từ tất cả phân xưởng theo trường Ngày.
2. Trong từng ngày, sắp xếp dữ liệu theo Thời gian, sau đó theo Tên phân xưởng.
3. Với mỗi dòng tổng hợp, tạo trường Nội dung báo cáo bằng cách ghép:
   - Tên trưởng ca
   - Nội dung công việc đạt được
4. Nếu thiếu Trưởng ca hoặc Nội dung, vẫn giữ dòng dữ liệu và ghi chú rõ phần bị thiếu.
5. Đánh lại STT của bảng tổng hợp từ 1 đến hết.
6. Trả về bảng tổng hợp có thể dùng trực tiếp cho bước 3 tạo file Excel tổng hợp.

Ràng buộc:
- Không tự ý loại bỏ dữ liệu trùng ngày, vì mỗi phân xưởng là một đơn vị báo cáo riêng.
- Không gộp nhiều phân xưởng vào cùng một ô nếu mục tiêu đầu ra là bảng Excel.
- Không sửa nội dung công việc gốc, chỉ được chuẩn hóa cách trình bày.
- Phải giữ được thông tin ngày, thời gian, tên phân xưởng và nội dung báo cáo.
- Nếu dữ liệu thiếu, dùng ghi chú ngắn như "[Thiếu trưởng ca]" hoặc "[Thiếu nội dung]".
- Định dạng ngày thống nhất dạng YYYY-MM-DD.
- Định dạng thời gian thống nhất dạng HH:MM:SS.

Định dạng đầu ra:
Bảng tổng hợp gồm các cột:
- STT: Số thứ tự sau khi tổng hợp
- Ngày: Ngày làm việc
- Thời gian: Thời gian bắt đầu ca
- Tên phân xưởng: Tên phân xưởng lấy từ tên file nguồn
- Nội dung: Tên trưởng ca và nội dung công việc đạt được

Quy tắc tạo cột Nội dung:
- Nếu có đủ Trưởng ca và Nội dung: "<Trưởng ca>: <Nội dung>"
- Nếu thiếu Trưởng ca: "[Thiếu trưởng ca]: <Nội dung>"
- Nếu thiếu Nội dung: "<Trưởng ca>: [Thiếu nội dung]"
- Nếu thiếu cả hai: "[Thiếu trưởng ca]: [Thiếu nội dung]"
```

## Thử nghiệm với dữ liệu hiện có

### Bảng tổng hợp mẫu

| STT | Ngày | Thời gian | Tên phân xưởng | Nội dung |
|---:|---|---|---|---|
| 1 | 2026-01-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn A: Báo cáo 1 |
| 2 | 2026-01-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn A: Báo cáo 1 |
| 3 | 2026-01-07 | 08:00:00 | Phân xưởng 3 | Trần Văn A: Báo cáo 1 |
| 4 | 2026-02-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn B: Báo cáo 2 |
| 5 | 2026-02-07 | 08:00:00 | Phân xưởng 2 | [Thiếu trưởng ca]: Báo cáo 2 |
| 6 | 2026-02-07 | 08:00:00 | Phân xưởng 3 | Trần Văn B: Báo cáo 2 |
| 7 | 2026-03-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn C: Báo cáo 3 |
| 8 | 2026-03-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn C: Báo cáo 3 |
| 9 | 2026-03-07 | 08:00:00 | Phân xưởng 3 | Trần Văn C: Báo cáo 3 |
| 10 | 2026-04-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn D: Báo cáo 4 |
| 11 | 2026-04-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn D: Báo cáo 4 |
| 12 | 2026-04-07 | 08:00:00 | Phân xưởng 3 | Trần Văn D: Báo cáo 4 |
| 13 | 2026-05-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn E: Báo cáo 5 |
| 14 | 2026-05-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn E: Báo cáo 5 |
| 15 | 2026-05-07 | 08:00:00 | Phân xưởng 3 | [Thiếu trưởng ca]: [Thiếu nội dung] |
| 16 | 2026-06-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn A: Báo cáo 6 |
| 17 | 2026-06-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn A: Báo cáo 6 |
| 18 | 2026-06-07 | 08:00:00 | Phân xưởng 3 | Trần Văn A: Báo cáo 6 |
| 19 | 2026-07-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn B: Báo cáo 7 |
| 20 | 2026-07-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn B: Báo cáo 7 |
| 21 | 2026-07-07 | 08:00:00 | Phân xưởng 3 | Trần Văn B: Báo cáo 7 |
| 22 | 2026-08-07 | 08:00:00 | Phân xưởng 1 | Nguyễn Văn C: Báo cáo 8 |
| 23 | 2026-08-07 | 08:00:00 | Phân xưởng 2 | Đoàn Văn C: Báo cáo 8 |
| 24 | 2026-08-07 | 08:00:00 | Phân xưởng 3 | Trần Văn C: Báo cáo 8 |

### Kết quả kiểm tra

| Nội dung kiểm tra | Kết quả |
|---|---:|
| Tổng số dòng đầu vào | 24 |
| Tổng số dòng sau tổng hợp | 24 |
| Số ngày làm việc | 8 |
| Số phân xưởng | 3 |
| Số dòng có cảnh báo thiếu dữ liệu | 2 |

### Cảnh báo còn tồn tại sau tổng hợp

| STT tổng hợp | Ngày | Tên phân xưởng | Vấn đề |
|---:|---|---|---|
| 5 | 2026-02-07 | Phân xưởng 2 | Thiếu trưởng ca |
| 15 | 2026-05-07 | Phân xưởng 3 | Thiếu trưởng ca và thiếu nội dung |

