# Context dự án EVNGENCO2 SCADA

## Phiên bản hiện tại

Phiên bản: `V1.1`

Script chính của phiên bản này: `main1_1.py`

## Mục tiêu

Dự án dùng để tổng hợp báo cáo công việc phân xưởng từ các file Excel trong thư mục `Data/`.

Quy trình xử lý gồm 3 bước:

1. Đọc toàn bộ file dữ liệu `.xlsx` trong `Data/`.
2. Tổng hợp nội dung công việc theo ngày từ tất cả phân xưởng.
3. Tạo file Excel tổng hợp theo cấu trúc báo cáo cuối cùng.

## Dữ liệu đầu vào

- Chỉ xử lý file `.xlsx` trong thư mục `Data/`.
- Bỏ qua file tạm Excel có tên bắt đầu bằng `~$`.
- Tên file như `PX_1.xlsx` được hiểu là `Phân xưởng 1`.
- Cột `Giờ` được chuẩn hóa thành `Thời gian`.
- Các trường quan trọng cần kiểm tra thiếu dữ liệu:
  - `Ngày`
  - `Thời gian`
  - `Trưởng ca`
  - `Nội dung`

## Kết quả đầu ra

`main1_1.py` tạo file Excel tổng hợp ở thư mục gốc dự án.

Tên file có dạng:

```text
Baocao_TH_<ngaygio>.xlsx
```

File Excel gồm các sheet:

- `Tong hop`: bảng tổng hợp cuối cùng.
- `Buoc 1`: thông tin file nguồn đã đọc.
- `Cau truc cot`: cấu trúc cột phát hiện và chuẩn hóa.
- `Canh bao du lieu`: danh sách cảnh báo dữ liệu thiếu.

Toàn bộ workbook được định dạng:

- Font: `Times New Roman`
- Cỡ chữ: `12`

## Tính năng log file

Phiên bản `V1.1` bổ sung tính năng ghi log file trong `main1_1.py`.

Mỗi lần chạy sẽ tạo một file log trong thư mục `logs/`.

Tên log có dạng:

```text
Baocao_TH_<ngaygio>.log
```

Log ghi lại:

- thời điểm bắt đầu và kết thúc quy trình;
- số file `.xlsx` phát hiện;
- từng file, sheet và phân xưởng đã đọc;
- cảnh báo dữ liệu thiếu;
- số dòng sau tổng hợp;
- tên file Excel đầu ra;
- lỗi chi tiết nếu quy trình bị lỗi.

## Cách chạy

```powershell
python main1_1.py
```

Sau khi chạy, kiểm tra:

- file Excel tổng hợp ở thư mục gốc dự án;
- file log tương ứng trong thư mục `logs/`.
