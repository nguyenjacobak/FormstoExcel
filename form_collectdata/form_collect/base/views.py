from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook,load_workbook
import os
import pandas as pd
from django.http import JsonResponse
import json

# Create your views here.
def get_students_view(request):
    data = request.body.decode('utf-8')
    project_name = json.loads(data).get('project_name', '')
    students = get_students_by_project_name(project_name)
    return JsonResponse(students, safe=False)

def get_all_students_view(request):
    students = get_all_students()
    return JsonResponse(students, safe=False)

def get_all_councils_view(request):
    councils = get_all_councils()
    return JsonResponse(councils, safe=False)

def get_students_by_project_name(project_name):
    print(project_name, "AAA")
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


def get_projects_by_lecture_and_type(lecturer_name, project_type):
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

####

#### Process all data
def get_all_councils():
    file_path = os.path.join('DataBase', 'db.xlsx')
    df = pd.read_excel(file_path, sheet_name="DS Hội đồng_DN", skiprows=2)
    df = df.dropna(ignore_index=True)
    df = df[df['Họ và tên'] != 'Họ và tên']
    df

    councils = {}
    for i in range(len(df)):
        member = df.iloc[i]
        name = member['Họ và tên']
        role = member['Nhiệm vụ'].replace('\xa0', '')
        unit = member['Đơn vị']
        id_council = f"HD{(i)//5 + 1}"
        if councils.get(id_council) is None:
            councils[id_council] = [{
                'name': name,
                'role': role,
                'unit': unit
            }]
        else:
            councils[id_council].append({
                'name': name,
                'role': role,
                'unit': unit
            })
    return councils

def get_all_students():
    # file_path = os.path.join('DataBase', 'db.xlsx')
    # df2 = pd.read_excel(file_path, sheet_name="DS SV các Hội đồng phân PB", skiprows=1)
    # df2 = df2.drop(columns = ['Unnamed: 0'])
    # df2.loc[1, 'Unnamed: 11'] = df2.loc[0, 'Unnamed: 11']
    # df2.loc[1, 'Unnamed: 12'] = df2.loc[0, 'Unnamed: 12']
    # df2.loc[1, 'Unnamed: 3'] = 'Tên'
    # df2.columns = df2.iloc[1]
    # df2 = df2.dropna(how='all', ignore_index=True)

    # councils = get_all_councils()
    # id_council = 'HD1'
    # students = {
    # }
    # for i in range(len(df2)):
    #     msv = df2.loc[i, 'Mã sinh viên']
    #     if 'Hội đồng' in msv:
    #         tmp = msv.split(' ')
    #         id_council = f"HD{tmp[-1]}"
    #         continue
    #     if msv == 'Mã sinh viên':
    #         continue
    #     name = df2.loc[i, 'Họ và tên'] + ' ' + df2.loc[i, 'Tên']
    #     day_of_birth = df2.loc[i, 'Năm sinh']
    #     class_name = df2.loc[i, 'Lớp']
    #     instructor = df2.loc[i, 'Giáo viên hướng dẫn']
    #     subject = df2.loc[i, 'Bộ môn ']
    #     project = df2.loc[i, 'Tên đề tài đồ án/ khóa luận tốt nghiệp']
    #     project_type = df2.loc[i, 'Loại đồ án']
    #     group = id_council + ' - ' + str(df2.loc[i, 'Nhóm']) if not pd.isnull(df2.loc[i, 'Nhóm']) else ''
    #     opponent = df2.loc[i, 'Phản biện  (Khoa)'] if not pd.isna(df2.loc[i, 'Phản biện  (Khoa)']) else ''
    #     council = councils[id_council]
    #     students[msv] = {
    #         'name': name,
    #         'day_of_birth': day_of_birth,
    #         'class_name': class_name,
    #         'instructor': instructor,
    #         'subject': subject,
    #         'project': project,
    #         'project_type': project_type,
    #         'group': group,
    #         'opponent': opponent,
    #         'council': council
    #     }

    #     # Get students from final sheet
    # new_students = students.copy()
    # for msv in students.keys():
    #     student = students[msv]
    #     if student['project'] == 'None':
    #         continue
    #     project_name = student['project']
    #     students_same_project = get_students_by_project_name(project_name)
    #     for std in students_same_project:
    #         msv_std = std['msv']
    #         if msv_std not in students.keys():
    #             fullname = std['fullname']
    #             day_of_birth = std['day_of_birth']
    #             class_name = std['class']
                
    #             new_students[msv_std] = students[msv].copy()
    #             new_students[msv_std]['name'] = fullname
    #             new_students[msv_std]['day_of_birth'] = day_of_birth
    #             new_students[msv_std]['class_name'] = class_name
    #             new_students[msv_std]['msv'] = msv_std
    new_students = []
    file_path = os.path.join('DataBase', 'students.json')
    with open(file_path, 'r', encoding='utf-8') as f:
        new_students = json.load(f)
    return new_students
#### End process all data

def index(request):
    lecturers = get_lecturers()
    return render(request, 'index.html', {'lecturers': lecturers})

def hoiDongChuyenMon(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            students = list(students.values())
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
            students = list(students.values())
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
            students = list(students.values())
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
            students = list(students.values())
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
            students = list(students.values())
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


