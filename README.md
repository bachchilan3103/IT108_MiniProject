# **MINI PROJECT** 
 HỆ THỐNG QUẢN TRỊ HẠ TẦNG & NHẬT KÝ KỸ THUẬT
## Module 1 : Thiết lập "căn cứ số"
- CMD `New-Item -ItemType Directory -Path GreenTech-system , GreenTech-system/docs , GreenTech-system/scrips , GreenTech-system/data , GreenTech-system/backup`
  Bash `mdkdir -p greentech-system/{docs,scrips,data,backup}`

- cd "GreenTech-system"
  sau khi di chuyển thư mục chúng ta sẽ tạo file cấu hình ẩn .env chứa mật khẩu hệ thống
- echo "DB_PASSWORD=DANG_CAP_ANH_BAY" > .env
## Modul2 : *Robot lọc Log tự động (Advanced Terminal & Scripting)*
- @echo off
echo ====== >> ...\backup\bao-cao-su-co.txt
echo BAO CAO SU CO >> ...\backup\bao-cao-su-co.txt
echo ====== >> ... \backup\bao-cao-su-co.txt
- cat ..\data\log-he-thong.log | grep /i "CRITICAL" >> ..\backup\bao-cao-su-co.txt
  cat ..\data\log-he-thong.log | grep /i "ERROR" >> ..\backup\bao-cao-su-co.txt 
echo *XONG* 
## Module 3: **Tài liệu hướng dẫn vận hành (Markdown & Word)** 
# GREENTECH
## Hệ Thống Quản Trị Hạ Tầng và Nhật Ký Kỹ Thuật
### 1. Hướng dẫn chạy scrips tự động hóa
 '''CMD
 cd `scrips`
 auto_report.bat
### 2. Checklist kiểm tra thiết bị 
- [x] xác minh dung lượng đĩa trống trên máy chủ log
- [ ] chạy script `auto_report.bat` mỗi ngày để tổng hợp danh sách
### 3. Bảng thông số phần cứng
| mã linh kiện | Tên thiết bị | loại storage | đơn giá(VND) |
| :--- | --- | --- | ---: |
| CPU | intel core i5 | N/A | 5,000,000 |
| RAM | 8GB | N/A | 600 |
| SSD | 256GB | SSD | 1,200,000 |

#### 1. cấu trúc chuẩn bị tài liệu SRS 
 - # **Chương 1 : Giới Thiệu và Tổng Quan** ( heading 1 )
 * 1.1 Mục Tiêu : **quản lý danh sách phần cứng của 100 trạm cảm biến IoT ngoài trang trại** ( heading 2 )
 * 1.2 **Tư duy hệ thống DIKW và IPO** ( heading 2 )
   * 1.2.1 DATA ( dữ liệu ) : file log 500MB ( heading 3 )
   * 1.2.2 INFORMATION ( thông tin ) : các dòng thông tin quang trọng chứa chữ "CRITCAL ( heading 3)
   * 1.2.3 KNOWLEDGE ( tri thức ) : đánh giá sự ổn định thiết bị ( heading 3 )
   * 1.2.4 WISDOM ( trí tuệ ) : thay thế phần cứng hỏng ( heading 3 )
 - # **Chương 2 : Kiến trúc hệ thống và lưu trữ** ( heading 1 )
   * 2.1 kiến trúc lưu trữ cây thưc mục 3 cấp ( heading 2 )
   * 2.2 mô hình 4 trụ cột ( heading 2 )
      * 2.2.1 bảo mật bí mật ( heading 3 )
      * 2.2.2 tự động hóa an toàn ( heading 3 )
      * 2.2.3 chuẩn hóa tên ( heading 3 )
      * 2.2.4 sao lưu độc lập ( heading 3 )
  - # **Chương 3 : Hướng dẫn vận hành và tra cứu** ( heading 1 )
    * 3.1 Hướng dẫn chạy script ( heading 2 )
    * 3.2 checklist kiểm tra thiết bị ( heading 2 )
    * 3.3 Bảng thông số phần cứng ( heading 2 )

  Chèn Auto TOC ở trang đầu tiên : ấn vào `references` và chọn `Table of contents` tiếp tục chọn `automatic table 1`
  ## Module 4 : Đối soát tài nguyên ( Excel )
  =VLOOKUP(TRIM(B2), 'danh-muc-linh-kien'!$A$2:$D$11, 2, FALSE)