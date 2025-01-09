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
    file_path = os.path.join('DataBase', 'student_list.xlsx')
    df = pd.read_excel(file_path)
    df = df.fillna(method='ffill')

    students = []
    for _, row in df.iterrows():
        if row['Tên đề tài đồ án/ khóa luận tốt nghiệp'] == project_name:
            student_id = row['Mã sinh viên']
            student_name = row['Họ và tên']
            student_class = row['Lớp']
            students.append({
                'student_id': student_id,
                'student_name': student_name,
                'student_class': student_class
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
    context = {
        'name': request.GET.get('name'),
        'project_type': request.GET.get('project_type'),
        'project_list': request.GET.get('project_list'),
        'form_type': request.GET.get('form_type'),
        'students': json.loads(request.GET.get('students', '[]'))
    }
    return render(request, 'hoiDongChuyenMon.html', context)
def baoCaoTienDoL1(request):
    context = {
        'name': request.GET.get('name'),
        'project_type': request.GET.get('project_type'),
        'project_list': request.GET.get('project_list'),
        'form_type': request.GET.get('form_type'),
        'students': json.loads(request.GET.get('students', '[]'))
    }
    return render(request, 'baoCaoTienDoL1.html', context)
def baoCaoTienDoL2(request):
    context = {
        'name': request.GET.get('name'),
        'project_type': request.GET.get('project_type'),
        'project_list': request.GET.get('project_list'),
        'form_type': request.GET.get('form_type'),
        'students': json.loads(request.GET.get('students', '[]'))
    }
    return render(request, 'baoCaoTienDoL2.html', context)
def huongdan3(request):
    context = {
        'name': request.GET.get('name'),
        'project_type': request.GET.get('project_type'),
        'project_list': request.GET.get('project_list'),
        'form_type': request.GET.get('form_type'),
        'students': json.loads(request.GET.get('students', '[]'))
    }
    return render(request, 'huongdan3.html', context)
def canBoPhanBien(request):
    context = {
        'name': request.GET.get('name'),
        'project_type': request.GET.get('project_type'),
        'project_list': request.GET.get('project_list'),
        'form_type': request.GET.get('form_type'),
        'students': json.loads(request.GET.get('students', '[]'))
    }
    return render(request, 'canBoPhanBien.html', context)
