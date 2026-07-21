# Thư mục ảnh tĩnh

Đặt file quốc huy vào đây với tên `quochuy.png`:

    public/quochuy.png

Ứng dụng sẽ tự lấy file này cho toàn bộ vị trí hiển thị quốc huy
(trang đăng nhập, sidebar, thanh mobile, màn hình đang tải, watermark).

Định dạng chấp nhận, theo thứ tự ưu tiên:
`quochuy.png` → `quochuy.svg` → `quochuy.jpg` → `quochuy.webp`

Nếu không tìm thấy file nào, ứng dụng tự vẽ lại quốc huy bằng SVG
(hàm `EmblemSvg` trong `src/App.jsx`) nên giao diện không bị vỡ.

Khuyến nghị: ảnh PNG **nền trong suốt**, chiều rộng khoảng 400–600px.
Nền trắng cũng dùng được nhưng sẽ lộ ô vuông trắng trên các nền màu.

Có thể xóa file README này sau khi đã bỏ ảnh vào.
