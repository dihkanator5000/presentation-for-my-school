# Bài trình bày: Phát triển mối quan hệ hòa đồng, hợp tác với thầy cô và bạn bè

Bài thuyết trình môn **Hoạt động trải nghiệm, hướng nghiệp – THCS** (18 slide, 2 tiết), tiếng Việt, do **học sinh thuyết trình** (xưng "em / chúng em").

## File PowerPoint
- **`Phát triển mối quan hệ hòa đồng, hợp tác - HĐTN lớp 7 (học sinh thuyết trình).pptx`** — file chính (tên có dấu).

Mở bằng Microsoft PowerPoint / Google Slides / LibreOffice; bấm **F5** để trình chiếu.

## Thiết kế
- **Giao diện SÁNG**, trọng tâm là **hiệu ứng kính lỏng (Liquid Glass)** kiểu iOS 26: thẻ kính mờ, nền gradient nhạt với các đốm màu mềm, viền sáng, bóng đổ mềm, vệt sáng trên đỉnh thẻ.
- Cảm giác **iPhone cao cấp**: bố cục rộng, chữ to đậm, chip/pill nhã nhặn, màu theo bảng màu hệ thống iOS.
- **Icon Apple SF Symbols** (không dùng emoji): file PNG nền trong suốt được **tải trực tiếp từ web** — bộ [SF Symbols Online](https://github.com/andrewtavis/sf-symbols-online/tree/master/glyphs). Không dùng Lucide, không tự vẽ, không render/tạo lại icon. Mỗi icon được giữ nguyên tỉ lệ và đặt cách chữ ít nhất một khoảng đệm rõ ràng.
- **Font có italic**: mặc định `Segoe UI` (đủ Regular/Italic/Bold). Nếu muốn đẹp hơn nữa, cài các font có sẵn khối italic và mở file chỉnh: `SF Pro Display`, `Be Vietnam Pro`, `Inter`, `Manrope` (đều hỗ trợ tiếng Việt + italic).
- **Animation mượt**: mỗi slide có chuyển cảnh **Fade**; khi trình chiếu bấm phím cách / chuột để nội dung **hiện dần mềm mại**.

## Nội dung
1. Bìa · Nội dung buổi trình bày · Mong muốn
2. Tiết 1 – Khám phá & chia sẻ: hòa đồng/hợp tác là gì, 4 "chìa khóa" kết nối, 4 bước làm việc nhóm, hợp tác với thầy cô, câu chuyện kể mẫu
3. Tiết 2 – Cùng thực hành: 3 tình huống SGK (việc riêng trong giờ thực hành · bạn ốm trước buổi biểu diễn · bạn mới chuyển đến nhút nhát), kịch bản phân vai mẫu, bài học rút ra, thông điệp, cam kết tuần, cảm ơn

Mỗi slide đều có **ghi chú (notes)** – lời thoại gợi ý để học sinh nói khi đứng lớp (mở tab View → Notes). Nhớ thay tên/lớp ở chỗ `[…]` trên slide 1 và slide 18.

## Tự dựng lại file (tuỳ chọn)
```bash
pip install python-pptx pillow
python3 build_liquid.py   # tạo file .pptx
```
- Icon PNG Apple SF Symbols đã tải nằm trong `assets/apple-sf/`; mã nguồn chỉ chèn các file này vào PowerPoint, không sinh icon.
- Các bản trình bày cũ (nền tối, chất Xbox, emoji) đã được **xoá khỏi repo** theo yêu cầu.
