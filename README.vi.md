# Exam Shuffler – Trộn đề thi tiếng Việt

Ứng dụng chạy trên máy cá nhân, đọc đề thi Word `.docx` và tạo nhiều mã đề. Giao diện, hướng dẫn, cảnh báo, kết quả kiểm tra và bảng đáp án sử dụng tiếng Việt. Tên tệp và nội dung có dấu tiếng Việt được hỗ trợ; nội dung đề gốc không được dịch hay dựng lại thành văn bản thuần.

## Khởi động trên Windows

Mở PowerShell trong thư mục dự án. Nếu môi trường `.venv` đã được cài đặt, chỉ cần chạy:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Cài đặt lần đầu với Python 3.11 trở lên:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Không bắt buộc kích hoạt môi trường ảo. Nếu đã kích hoạt, có thể dùng `python -m streamlit run app.py`.

## Khởi động trên Ubuntu

Sau khi cài Python 3.11 trở lên và gói `python3-venv`:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Mở **http://127.0.0.1:8501** trên trình duyệt. Nhấn Ctrl+C trong cửa sổ chạy ứng dụng để dừng. Ứng dụng mặc định chỉ truy cập được từ máy của bạn; tài liệu được xử lý tại máy và không gửi đến API bên ngoài.

## Các bước sử dụng

1. **Tải đề thi lên:** chọn tệp `.docx`, tối đa 30 MiB.
2. **Phân tích và kiểm tra:** mở từng phần để xem câu hỏi, phương án, tư liệu, bảng và điểm. Đánh dấu **Tôi đã kiểm tra các câu hỏi được nhận diện và các cảnh báo**. Lỗi nghiêm trọng cần được sửa trong tệp gốc trước khi tạo đề.
3. **Thiết lập trộn đề:** mặc định tạo 4 mã đề từ 101 đến 104. Chọn trộn câu hỏi trong từng phần, phương án A–D, nhận định a–d hoặc câu tự luận. Trộn tự luận cần bật cả trộn câu hỏi. Cùng đề gốc, thiết lập và số khởi tạo ngẫu nhiên sẽ cho cùng kết quả.
4. **Nhập đáp án (tùy chọn):** nhập theo số câu và nhãn của đề gốc. Câu đúng/sai phải nhập đủ bốn nhận định hoặc để trống toàn bộ. Màu chữ chỉ là gợi ý; cần kiểm tra và xác nhận. Sửa đáp án sẽ hủy xác nhận trước đó để bạn kiểm tra lại.
5. **Tạo và tải đề:** nhấn **Tạo các mã đề**, chờ kiểm tra hoàn tất rồi tải từng đề, bảng đáp án hoặc toàn bộ ZIP. Khi thay đổi thiết lập hoặc đáp án, cần tạo lại đề để cập nhật tệp tải về.

Tên các tệp xuất vẫn là `Exam_101.docx`, `Exam_102.docx`, …, `Answer_Key.docx` và `Exam_Shuffler.zip`. Bảng đáp án ghi **Câu**, **Đáp án**, **Mã câu gốc** và các giá trị **Đúng/Sai** bằng tiếng Việt. Không nhập đáp án vẫn tạo được đề, nhưng không có bảng đáp án.

## Cấu trúc đề được hỗ trợ

- Nhãn câu hỏi: `Câu 1:`, `Câu 1.`, `Câu 1. (2,0 điểm).`, kể cả khi số câu nằm trên nhiều đoạn chữ có định dạng khác nhau.
- Tiêu đề phần: `PHẦN I.`, `PHẦN II.`, `PHẦN 1.`, … Không giới hạn đề phải có ba phần hay số câu cố định.
- Trắc nghiệm có bốn phương án `A.`–`D.`; đúng/sai có bốn nhận định `a)`–`d)`. Các nhãn trong đề gốc cần đủ và đúng thứ tự.
- Phương án nằm trên các đoạn riêng hoặc cùng một đoạn đơn giản, ngăn cách bằng khoảng trắng/tab.
- Tư liệu và bảng nằm trong nhóm câu hỏi được giữ cùng câu hỏi khi di chuyển. Điểm tự luận được giữ cùng nội dung câu tương ứng.
- Dấu kết thúc như `----------- HẾT ----------` giúp cố định phần cuối đề. Nếu thiếu, ứng dụng cảnh báo để kiểm tra câu cuối.

## Lưu ý về định dạng

Ứng dụng sao chép các khối OOXML gốc và chỉ sửa số câu/nhãn phương án. Các thành phần khác của DOCX, gồm ảnh, đầu trang, chân trang, liên kết và nội dung nhúng, giữ nguyên dữ liệu. Tệp gốc không bị ghi đè.

Kiểm tra cấu trúc không thay thế việc xem bố cục thực tế. Trộn câu và thêm mã đề có thể thay đổi cách xuống trang. Hãy mở từng tệp trong trình đọc Word để kiểm tra trước khi in.

- Câu hỏi bắt đầu trong bảng/hộp văn bản, thay đổi chưa chấp nhận, chữ ký số, đánh số tự động và ngắt phần nằm trong câu hỏi có thể bị chặn để tránh xuất sai cấu trúc.
- Với phương án có bố cục phức tạp, bảng sau nhãn phương án hoặc vùng đánh dấu kéo dài qua nhiều phương án, hãy tắt trộn phương án theo cảnh báo.
- Tư liệu dùng chung cho nhiều câu riêng biệt cần được kiểm tra thủ công; có thể tắt trộn câu hỏi để giữ thứ tự.
- Trường Word và số trang được giữ nhưng không tính lại; có thể cần cập nhật khi mở tài liệu.
- Màu chữ được giữ trong đề xuất. Nếu màu làm lộ đáp án, hãy bỏ màu trong bản đề dành cho học sinh trước khi tải lên.
- Xóa mã đề cố định cũ trong mẫu để tránh xung đột với mã đề mới.
- Một số chữ bên trong thành phần tải tệp hoặc trình đơn hệ thống của Streamlit có thể vẫn hiển thị bằng tiếng Anh.

Xem [README.md](README.md) để biết chi tiết kiến trúc và các giới hạn DOCX.
