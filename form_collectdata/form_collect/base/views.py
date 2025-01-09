from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook,load_workbook
import os
# Create your views here.
def get_lecturers():
    file_path = os.path.join('DataBase', 'db.xlsx')
    workbook = load_workbook(filename=file_path)
    sheet = workbook.worksheets[3]  # Trang thứ 3 của file Excel
    lecturers = set()

    for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):
        name = str(row[0])  # Cột A
        if name and '(' not in name and ')' not in name:
            lecturers.add(name)
        name = str(row[2])  # Cột C
        if name and '(' not in name and ')' not in name:
            lecturers.add(name)

    return lecturers
def get_projects(lecturer_name, project_type):
    file_path = os.path.join('DataBase', 'db.xlsx')
    workbook = load_workbook(filename=file_path)
    sheet = workbook.worksheets[4]  # Trang thứ 4 của file Excel
    projects = []

    # Tìm các cột dựa trên tiêu đề
    headers = {cell.value: idx for idx, cell in enumerate(sheet[1])}
    lecturer_col = headers.get('Giáo viên hướng dẫn')
    project_name_col = headers.get('Tên đề tài đồ án/ khóa luận tốt nghiệp')
    project_type_col = headers.get('Làm đồ án/Học phần TTTN')

    if lecturer_col is None or project_name_col is None or project_type_col is None:
        return projects  # Trả về danh sách rỗng nếu không tìm thấy các cột cần thiết

    for row in sheet.iter_rows(min_row=2, values_only=True):
        lecturer = str(row[lecturer_col]).split('.')[-1].strip()  # Lấy phần tên sau học vị
        if lecturer == lecturer_name:
            project_name = row[project_name_col]
            project_category = row[project_type_col]
            if (project_type == "Cá nhân" and project_category == "Đồ án cá nhân") or \
               (project_type == "Nhóm" and project_category == "Đồ án nhóm"):
                projects.append(project_name)

    return projects
def index(request):
    lecturers = get_lecturers()
    selected_lecturer = request.GET.get('name')
    selected_project_type = request.GET.get('project_type')
    projects = []

    if selected_lecturer and selected_project_type:
        projects = get_projects(selected_lecturer, selected_project_type)

    return render(request, 'index.html', {'lecturers': lecturers, 'projects': projects})
def form1(request):
    return render(request,'hoiDongChuyenMon.html')
def form2(request):
    return render(request,'canBoPhanBien.html')
def formhd3(request):
    return render(request,'huongdan3.html')
