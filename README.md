Web Chấm Đồ Án Tốt Nghiệp - HVKTBCVT

Mô tả dự án

Web chấm đồ án tốt nghiệp là một hệ thống giúp quản lý và đánh giá các bài đề tốt nghiệp dựa trên giao diện web. Hệ thống được phát triển bằng Django, đáp ứng nhu cầu quản lý học liệu và tổ chức chấm bài một cách hiệu quả.

Hướng dẫn cài đặt

Yêu cầu hệ thống

Python phiên bản >= 3.9

pip (Python Package Installer)

Các bước cài đặt

Clone repository

Sử dụng lệnh sau để clone repository từ GitHub:

git clone <https://github.com/nguyenjacobak/FormstoExcel>

cd Lab_project\Web_form_collect_data\FormstoExcel\form_collectdata\form_collect

Tạo virtual environment

Tạo môi trường ảo để quản lý các thư viện:

python -m venv venv
source venv/bin/activate      # Trên macOS/Linux
venv\Scripts\activate       # Trên Windows

Cài đặt các thư viện cần thiết

Chạy lệnh sau để cài đặt các thư viện:

pip install django
pip install openpyxl
pip install pandas
pip install numpy

Chạy lệnh:

python manage.py migrate

Chạy server

Chạy lệnh:

python manage.py runserver

Các chức năng chính

Quản lý danh sách sinh viên và giáo viên.

Tải lên và chấm điểm các bài đề tốt nghiệp.

Xuất báo cáo dưới dạng Excel.

Hướng dẫn :

-File final_new.xlsx chứa dữ liệu tổng mà các giảng viên đã chấm và submit

-File TongHopDiem1.xlsx là bảng điểm tổng hợp số 4 trong mẫu chấm

-File TongHopDiem2.xlsx là bảng điểm tổng hợp số 5 trong mẫu chấm

-Sau mỗi lần submit phải mở lại các file excel để dữ liệu được cập nhật

*LƯU Ý:

-Không được sửa, xóa hay thêm các cột vào cấc file excel đã tổng hợp

-Không di chuyển hay xóa các file trong các thư mục
