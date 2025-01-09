from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook,load_workbook
import os
import pandas as pd
from django.http import JsonResponse
import json

# Create your views here.
def get_students_view(request):
    project_name = request.GET.get('project_name')
    students = get_students(project_name)
    return JsonResponse(students, safe=False)

def get_students(project_name):
    file_path = os.path.join('DataBase', 'db.xlsx')
    df = pd.read_excel(file_path, sheet_name="Danh sách các đồ án ", skiprows=12)
    df = df.fillna(method='ffill')

    students = []
    col = 'Tên đề tài đồ án/ khóa luận tốt nghiệp'
    df_students = df[df[col].str.contains(project_name, case=False)]
    
    first_students = list(df_students['Họ và tên'])
    second_students = list(df_students['Unnamed: 3'])
    fullname_students = [f"{first} {second}" for first, second in zip(first_students, second_students)]
    msv_students = list(df_students['Mã sinh viên'])
    day_of_birth_students = list(df_students['Năm sinh'])
    class_students = list(df_students['Lớp'])
    for i in range(len(fullname_students)):
        students.append({
            'fullname': fullname_students[i],
            'msv': msv_students[i],
            'day_of_birth': day_of_birth_students[i],
            'class': class_students[i]
        })
    return students

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
        if name and '(' not in name and ')' not in name and name != 'None':
            lecturers.add(name)

    return [lecture for lecture in lecturers if lecture is not None]


def get_projects(lecturer_name, project_type):
    file_path = os.path.join('DataBase', 'db.xlsx')
    df = pd.read_excel(file_path, sheet_name="Danh sách các đồ án ", skiprows=12)
    df = df.fillna(method='ffill')

    projects = []
    cols = ['Tên đề tài đồ án/ khóa luận tốt nghiệp', 'Giáo viên hướng dẫn', 'Làm đồ án/Học phần TTTN']
    for col in cols:
        if col not in df.columns:
            return projects

    projects = df[df['Giáo viên hướng dẫn'].str.contains(lecturer_name, case=False) & df['Làm đồ án/Học phần TTTN'].str.contains(project_type, case=False)]['Tên đề tài đồ án/ khóa luận tốt nghiệp']

    projects = list(set(projects))
    return projects


def index(request):
    lecturers = get_lecturers()
    selected_lecturer = request.GET.get('name')
    selected_project_type = request.GET.get('project_type')
    projects = []

    if selected_lecturer and selected_project_type:
        projects = get_projects(selected_lecturer, selected_project_type)

    return render(request, 'index.html', {'lecturers': lecturers, 'projects': projects})

def hoiDongChuyenMon(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')
        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)

        context = {
            'students': students,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName
        }
        print(context)
        return render(request, 'hoiDongChuyenMon.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': ''
    }
    return render(request, 'hoiDongChuyenMon.html', context)

def baoCaoTienDoL1(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')
        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)

        context = {
            'students': students,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName
        }
        print(context)
        return render(request, 'baoCaoTienDoL1.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': ''
    }
    return render(request, 'baoCaoTienDoL1.html', context)

def baoCaoTienDoL2(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')
        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)

        context = {
            'students': students,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName
        }
        print(context)
        return render(request, 'baoCaoTienDoL2.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': ''
    }
    return render(request, 'baoCaoTienDoL2.html', context)

def huongdan3(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')
        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)

        context = {
            'students': students,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName
        }
        print(context)
        return render(request, 'huongdan3.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': ''
    }
    return render(request, 'huongdan3.html', context)

def canBoPhanBien(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')
        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)

        context = {
            'students': students,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName
        }
        print(context)
        return render(request, 'canBoPhanBien.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': ''
    }
    return render(request, 'canBoPhanBien.html', context)
