# CHƯƠNG TRÌNH ĐIỀN SỐ LIỆU VÀO FORM 

Thư mục chương trình: **Excel2ctt\bin\Debug**\
Chương trình có thể chạy ngay, không cần cài đặt

Xem thêm video hoạt động\
[![Xem Video](F.OkayToStart.png)](https://player.vimeo.com/video/1007337512?badge=0&amp;autopause=0&amp;player_id=0&amp;app_id=58479)

## HƯỚNG DẪN SỬ DỤNG

0. Lựa chọn template sẵn có phù hợp, kèm theo ảnh hướng dẫn trong mỗi thư mục. Copy tất cả các file trong đó, và ghi đè vào các file hiện có trong thư mục chứa file **Excel2ctt.exe**.\
   ![alt text](0.ChonTemplatePhuHop.png)\

1. Trích xuất danh sách cần nhập liệu. Ví dụ từ qldt.hust.edu.vn.
2. Đưa thông tin của SV vào file **diem.txt** (có sẵn ở bước trước đó) theo qui tắc:
   - Thông tin của của mỗi SV nằm trên một dòng.
   - Các trường thông tin của SV cách nhau bởi kí tự TAB.
3. Trên trình duyệt hoặc ứng dụng cần nhập liệu như MSTeams, mở tới form cần nhập liệu, _ví dụ giao diện thêm sinh viên vào lớp học_ \
   ![4.DatDauNhacPromtVaoTextBoxThemThanhVienTrongMSTeams.png](Excel2ctt\bin\Debug\DangKySVvaoTeam_1cot\4.DatDauNhacPromtVaoTextBoxThemThanhVienTrongMSTeams.png)
4. Bấm vào**Excel2ctt.exe**. Phần mềm sẽ hiện ra thông tin về dòng dầu tiên. Kiểm tra kĩ xem có bị nhầm nội dung không. Bấm **OK**.\
   ![Bấm okay để bắt đầu](F.OkayToStart.png)]
5. Focus (Quay trở lại)  trình duyệt hoặc ứng dụng cần nhập liệu như MSTeams ngay, bảo đảm rằng dấu nhắc promt vẫn ở ô nhập liệu đầu tiên.\
   Đợi vài giây là xong. Chương trình sẽ bắt đầu điền số liệu

## HƯỚNG DẪN NÂNG CAO

Phần mềm cho phép tinh chỉnh một số cấu hình chạy. Các cấu 
hình này được lưu và chỉnh sửa trong file Excel2ctt.exe.config
ở cùng thư mục chạy. Cụ thể là: 

GOTONEXTRECORD: Các kí tự được gửi đi sau khi kết thúc một dòng 
nhập liệu, để chuyển sang dòng kế tiếp. Mặc định là {TAB}{TAB}

WAITFORSWITCHINGAPP: Thời gian đợi để chuyển từ app này sang
website của phòng đào tạo. Đơn vị giây. Mặc định 5 giây.

## MÃ NGUỒN
------------------------------------------------------------
Mã nguồn mở tại : https://github.com/neittien0110/Excel2ctt
Mục đích công bố mã nguồn mở để bảo đảm tính trung thực của
chương trình, không có bất cứ tác động nào làm thay đổi giá
trị đầu vào.