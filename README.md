# Bài trình bày: Phát triển mối quan hệ hòa đồng, hợp tác với thầy cô và bạn bè

Bài thuyết trình môn **Hoạt động trải nghiệm, hướng nghiệp – THCS** (18 slide, 2 tiết), tiếng Việt, do **học sinh thuyết trình** (xưng "em / chúng em").

## File PowerPoint
- **`Phát triển mối quan hệ hòa đồng, hợp tác - HĐTN lớp 7 (học sinh thuyết trình).pptx`** — file chính (tên có dấu).

Bản hiện tại được ưu tiên cho **LibreOffice Impress**; mở file rồi bấm **F5** để trình chiếu. PowerPoint và Google Slides vẫn mở được.

## Thiết kế
- **Giao diện SÁNG**, trọng tâm là **hiệu ứng kính lỏng (Liquid Glass)** kiểu iOS 26: nền gradient nhạt với các đốm màu mềm, viền sáng và thẻ kính trắng có độ trong vừa phải.
- Bản LibreOffice-safe: các thẻ được làm sáng/đậm hơn, không dựa vào bóng mờ hoặc hiệu ứng trong suốt quá mạnh để chữ không bị chìm hay lệch.
- Cảm giác **iPhone cao cấp**: bố cục rộng, chữ to đậm, chip/pill nhã nhặn, màu theo bảng màu hệ thống iOS.
- **Icon Apple SF Symbols** (không dùng emoji): file PNG nền trong suốt được **tải trực tiếp từ web** — bộ [SF Symbols Online](https://github.com/andrewtavis/sf-symbols-online/tree/master/glyphs). Không dùng Lucide, không tự vẽ, không render/tạo lại icon. Mỗi icon được giữ nguyên tỉ lệ và đặt cách chữ ít nhất một khoảng đệm rõ ràng.
- **Font có italic**: mặc định `Liberation Sans` — font LibreOffice dùng ổn định, có Regular/Italic/Bold và hỗ trợ tiếng Việt. Không cần cài `Segoe UI` để giữ bố cục.
- **Chuyển động an toàn cho LibreOffice**: mỗi slide có chuyển cảnh **Fade**. Nội dung hiện đầy đủ ngay khi tới slide (không dùng animation riêng từng icon/chữ vì Impress có thể làm chúng ẩn hoặc lệch).

## Nội dung
1. Bìa · Nội dung buổi trình bày · Mong muốn
2. Tiết 1 – Khám phá & chia sẻ: hòa đồng/hợp tác là gì, 4 "chìa khóa" kết nối, 4 bước làm việc nhóm, hợp tác với thầy cô, câu chuyện kể mẫu
3. Tiết 2 – Cùng thực hành: 3 tình huống SGK (việc riêng trong giờ thực hành · bạn ốm trước buổi biểu diễn · bạn mới chuyển đến nhút nhát). **Nhóm em lần lượt đặt 3 câu hỏi mở, cả lớp giơ tay trả lời, sau đó nhóm em tổng hợp và chốt kĩ năng**; kèm kịch bản hỏi–đáp mẫu, bài học rút ra, thông điệp, cam kết tuần, cảm ơn.

Mỗi slide đều có **ghi chú (notes)** – lời thoại gợi ý để học sinh nói khi đứng lớp (mở tab View → Notes). Nhớ thay tên/lớp ở chỗ `[…]` trên slide 1 và slide 18.

## Đồng hồ suy nghĩ ở phần tình huống
Slide 11–13 có đồng hồ **`00:30` đếm ngược thật**, kích thước vừa phải ở góc phải trên.

- Đồng hồ **không tự chạy** khi tới slide; bấm trực tiếp vào vùng `00:30` để bắt đầu.
- Trong LibreOffice Impress, bấm lại video để Pause hoặc bấm **`DỪNG · RESET`** để quay lại trạng thái `00:30`.
- Đồng hồ được nhúng dạng video MP4 H.264 (có poster tĩnh), thay vì GIF tự chạy hay animation PowerPoint.

Muốn đổi thời gian, sửa một dòng trong `build_liquid.py` rồi dựng lại file:
```python
THINK_SECONDS = 30   # ví dụ: 15, 45 hoặc 60
```

## Tự dựng lại file (tuỳ chọn)
```bash
pip install python-pptx pillow imageio-ffmpeg
python3 build_liquid.py   # tạo file .pptx và video đồng hồ
```
- Icon PNG Apple SF Symbols đã tải nằm trong `assets/apple-sf/`; mã nguồn chỉ chèn các file này vào PowerPoint, không sinh icon.
- Các bản trình bày cũ (nền tối, chất Xbox, emoji) đã được **xoá khỏi repo** theo yêu cầu.
