from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook,load_workbook
import os
import pandas as pd
from django.http import JsonResponse
import json
from django.shortcuts import redirect

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
    lecturesList = [lecture for lecture in lecturers if lecture is not None]
    return sorted(lecturesList, key=lambda lecture: lecture.split()[-1])


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

def find_student_by_council_and_group_id(council_id=None, group_id=None):
    students = []
    if group_id is None or council_id is None:
        return students
    file_path = os.path.join('DataBase', 'db.xlsx')
    df2 = pd.read_excel(file_path, sheet_name="DS SV các Hội đồng phân PB", skiprows=1)
    df2 = df2.drop(columns = ['Unnamed: 0'])
    df2.loc[1, 'Unnamed: 11'] = df2.loc[0, 'Unnamed: 11']
    df2.loc[1, 'Unnamed: 12'] = df2.loc[0, 'Unnamed: 12']
    df2.loc[1, 'Unnamed: 3'] = 'Tên'
    df2.columns = df2.iloc[1]
    df2 = df2.dropna(how='all', ignore_index=True)
    index_council = find_index_hd_in_excel(df2, council_id)
    index_council_next = find_index_hd_in_excel(df2, council_id + 1)
    if index_council_next == -1:
        index_council_next = len(df2)
    df2 = df2.iloc[index_council+2:index_council_next]
    df2 = df2.reset_index(drop=True)
    for i in range(len(df2)):
        df2 = df2.reset_index(drop=True) 
        if df2.loc[i, 'Nhóm'] == group_id:
            students.append(df2.loc[i, 'Mã sinh viên'])
    return students

def find_index_hd_in_excel(df2, id_hd):
    for i in range(len(df2)):
        if df2.loc[i, 'Mã sinh viên'] == f'Hội đồng {id_hd}':
            return i
    return -1
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
        studentsSameGroup = []
        try:
            msv_list = request.GET.getlist('msv', '')
            council_id, group_id = None, None
            group = request.GET.get('group', None)
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            students = data.get('students', [])
            data = data.get('data', {})
            studentsData = [students[msv] for msv in msv_list]
            name = data.get('name', '')
            project_type = data.get('projectType', '')
            projectName = data.get('projectName', '')

            studentsMsvSameGroup = []
            if group is not None and group != '':
                tmp = group.split(' - ')
                council_id = int(tmp[0][2:])
                group_id = int(tmp[1])
                print(council_id, group_id)
                studentsMsvSameGroup = find_student_by_council_and_group_id(council_id, group_id)
                print(studentsMsvSameGroup)
            studentsSameGroup = [students[msv] for msv in studentsMsvSameGroup]
            


        except (json.JSONDecodeError, AttributeError):
            return JsonResponse({'error': 'Dữ liệu không hợp lệ hoặc trống'}, status=400)
        except Exception as e:
            print(e)

        context = {
            'students': studentsData,
            'students_count': len(students),
            'name': name,
            'project_type': project_type,
            'project_name': projectName,
            'studentsSameGroup': studentsSameGroup
        }
        print(context)
        return render(request, 'hoiDongChuyenMon.html', context)
    
    # Xử lý GET request (truy cập trực tiếp từ trình duyệt)
    context = {
        'students': [],
        'students_count': 0,
        'name': '',
        'project_type': '',
        'project_name': '',
        'studentsSameGroup': []
    }
    return render(request, 'hoiDongChuyenMon.html', context)

def baoCaoTienDoL1(request):
    if request.method == 'POST':
        try:
            msv_list = request.GET.getlist('msv', '')
            data = json.loads(request.body.decode('utf-8'))  # Lấy dữ liệu từ body
            data = data.get('data', {})
            students = data.get('students', [])
            students = [students[msv] for msv in msv_list]
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

def process_form_hd1(request):
    if request.method == 'POST':
        # Lấy số lượng sinh viên
        try:
            students_count = int(request.POST.get('students_count', 0))
        except ValueError:
            students_count = 0

        # Lấy dữ liệu từ form
        students = []
        for i in range(1, students_count + 1):
            student = {
                'fullname': request.POST.get(f'student_fullname_{i}', '').strip(),
                'msv': request.POST.get(f'student_msv_{i}', '').strip(),
                'class': request.POST.get(f'student_class_{i}', '').strip(),
                'diemC11': request.POST.get(f'diemC11SV{i}', '').strip(),
                'diemC12': request.POST.get(f'diemC12SV{i}', '').strip(),
                'diemC51': request.POST.get(f'diemC51SV{i}', '').strip(),
                'gpa': request.POST.get(f'gpaSV{i}', '').strip()
            }
            if student['msv']:  # Đảm bảo có mã sinh viên
                students.append(student)

        nhanXet = request.POST.get('nhanXet', '').strip()
        lecturer_name = request.POST.get('lecturer_name', '').strip()
        project_type = request.POST.get('project_type', '').strip()
        project_name = request.POST.get('project_name', '').strip()
        form_type = 'hd'

        # Đường dẫn tới thư mục và file Excel
        data_dir = 'DataCollected'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        file_path = os.path.join(data_dir, 'data.xlsx')

        # Mở hoặc tạo mới file Excel
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            workbook.save(file_path)

        # Mapping headers to column indexes
        headers = {}
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, cell in enumerate(header_row):
            if cell:
                headers[cell.strip().lower()] = idx + 1  # 1-based index

        required_headers = ["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5","hd.01", "hd.02", "hd.03","pb"]
        # Kiểm tra các header cần thiết
        missing_headers = [h for h in required_headers if h not in headers]
        if missing_headers:
            # Nếu thiếu header, thêm lại header chuẩn
            sheet.delete_rows(1)
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            
            workbook.save(file_path)
            headers = {cell.strip().lower(): idx +1 for idx, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))}

        # Duyệt qua từng sinh viên và lưu thông tin vào file Excel
        for student in students:
            msv = student['msv']
            # Tìm sinh viên theo msv ở cột 'msv'
            student_found = False
            for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
                cell_msv = row[headers['msv'] -1]
                if cell_msv.value == msv:
                    student_found = True
                    # Tìm cột hd.01 đến hd.03 trống
                    for hd_col in ["hd.01", "hd.02", "hd.03"]:
                        cell_hd = row[headers[hd_col] -1]
                        if not cell_hd.value:
                            cell_hd.value = f"{lecturer_name} - C1.1: {student['diemC11']} - C1.2: {student['diemC12']} - C5.1: {student['diemC51']} - GPA: {student['gpa']}"
                            break  # Ghi xong vào cột trống, dừng tìm cột tiếp theo
                    break  # Đã tìm thấy sinh viên, dừng tìm kiếm

            if not student_found:
                # Thêm sinh viên mới và ghi vào hd.01
                new_row = [
                    student['fullname'],
                    student['msv'],
                    student['class'],
                    "",
                    "",
                    "",
                    "",
                    "",
                    f"{lecturer_name} - C1.1: {student['diemC11']} - C1.2: {student['diemC12']} - C5.1: {student['diemC51']} - GPA: {student['gpa']}",
                    "",
                    "",
                    ""
                ]
                sheet.append(new_row)
                print(f"Thêm sinh viên mới: {msv} và ghi dữ liệu vào hd.01")

        # Lưu file Excel
        workbook.save(file_path)

        # Chuyển hướng đến trang testOutput.html với dữ liệu
        return render(request, 'testOutput.html', {
            'students': students,
            'nhanXet': nhanXet,
            'lecturer_name': lecturer_name,
            'project_type': project_type,
            'project_name': project_name,
            'form_type': form_type
        })

    return redirect('baoCaoTienDoL1')

def process_form_hd2(request):
    if request.method == 'POST':
        # Lấy số lượng sinh viên
        try:
            students_count = int(request.POST.get('students_count', 0))
        except ValueError:
            students_count = 0

        # Lấy dữ liệu từ form
        students = []
        for i in range(1, students_count + 1):
            student = {
                'fullname': request.POST.get(f'student_fullname_{i}', '').strip(),
                'msv': request.POST.get(f'student_msv_{i}', '').strip(),
                'class': request.POST.get(f'student_class_{i}', '').strip(),
                'diemC21': request.POST.get(f'diemC21SV{i}', '').strip(),
                'diemC22': request.POST.get(f'diemC22SV{i}', '').strip(),
                'diemC31': request.POST.get(f'diemC31SV{i}', '').strip(),
                'diemC52': request.POST.get(f'diemC52SV{i}', '').strip(),
                'gpa': request.POST.get(f'gpaSV{i}', '').strip()
            }
            if student['msv']:  # Đảm bảo có mã sinh viên
                students.append(student)

        nhanXet = request.POST.get('nhanXet', '').strip()
        lecturer_name = request.POST.get('lecturer_name', '').strip()
        project_type = request.POST.get('project_type', '').strip()
        project_name = request.POST.get('project_name', '').strip()
        form_type = 'hd'

        # Đường dẫn tới thư mục và file Excel
        data_dir = 'DataCollected'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        file_path = os.path.join(data_dir, 'data.xlsx')

        # Mở hoặc tạo mới file Excel
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            workbook.save(file_path)

        # Mapping headers to column indexes
        headers = {}
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, cell in enumerate(header_row):
            if cell:
                headers[cell.strip().lower()] = idx + 1  # 1-based index

        required_headers = ["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5","hd.01", "hd.02", "hd.03","pb"]
        # Kiểm tra các header cần thiết
        missing_headers = [h for h in required_headers if h not in headers]
        if missing_headers:
            # Nếu thiếu header, thêm lại header chuẩn
            sheet.delete_rows(1)
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            
            workbook.save(file_path)
            headers = {cell.strip().lower(): idx +1 for idx, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))}

        # Duyệt qua từng sinh viên và lưu thông tin vào file Excel
        for student in students:
            msv = student['msv']
            # Tìm sinh viên theo msv ở cột 'msv'
            student_found = False
            for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
                cell_msv = row[headers['msv'] -1]
                if cell_msv.value == msv:
                    student_found = True
                    # Tìm cột hd.01 đến hd.03 trống
                    for hd_col in ["hd.01", "hd.02", "hd.03"]:
                        cell_hd = row[headers[hd_col] -1]
                        if not cell_hd.value:
                            cell_hd.value = f"{lecturer_name} - C2.1: {student['diemC21']} - C2.2: {student['diemC22']} - C3.1: {student['diemC31']} - C5.2: {student['diemC52']} - GPA: {student['gpa']}"
                            break  # Ghi xong vào cột trống, dừng tìm cột tiếp theo
                    break  # Đã tìm thấy sinh viên, dừng tìm kiếm

            if not student_found:
                # Thêm sinh viên mới và ghi vào hd.01
                new_row = [
                    student['fullname'],
                    student['msv'],
                    student['class'],
                    "",
                    "",
                    "",
                    "",
                    "",
                    f"{lecturer_name} - C2.1: {student['diemC21']} - C2.2: {student['diemC22']} - C3.1: {student['diemC31']} - C5.2: {student['diemC52']} - GPA: {student['gpa']}",
                    "",
                    "",
                    ""
                ]
                sheet.append(new_row)
                print(f"Thêm sinh viên mới: {msv} và ghi dữ liệu vào hd.01")

        # Lưu file Excel
        workbook.save(file_path)

        # Chuyển hướng đến trang testOutput.html với dữ liệu
        return render(request, 'testOutput.html', {
            'students': students,
            'nhanXet': nhanXet,
            'lecturer_name': lecturer_name,
            'project_type': project_type,
            'project_name': project_name,
            'form_type': form_type
        })

    return redirect('baoCaoTienDoL2')

def process_form_hd3(request):
    if request.method == 'POST':
        # Lấy số lượng sinh viên
        try:
            students_count = int(request.POST.get('students_count', 0))
        except ValueError:
            students_count = 0

        # Lấy dữ liệu từ form
        students = []
        for i in range(1, students_count + 1):
            student = {
                'fullname': request.POST.get(f'student_fullname_{i}', '').strip(),
                'msv': request.POST.get(f'student_msv_{i}', '').strip(),
                'class': request.POST.get(f'student_class_{i}', '').strip(),
                'diemC23': request.POST.get(f'diemC23SV{i}', '').strip(),
                'diemC32': request.POST.get(f'diemC32SV{i}', '').strip(),
                'diemC41': request.POST.get(f'diemC41SV{i}', '').strip(),
                'diemC61': request.POST.get(f'diemC61SV{i}', '').strip(),
                'diemC62': request.POST.get(f'diemC62SV{i}', '').strip(),
                'gpa': request.POST.get(f'gpaSV{i}', '').strip()
            }
            if student['msv']:  # Đảm bảo có mã sinh viên
                students.append(student)

        nhanXet = request.POST.get('nhanXet', '').strip()
        lecturer_name = request.POST.get('lecturer_name', '').strip()
        project_type = request.POST.get('project_type', '').strip()
        project_name = request.POST.get('project_name', '').strip()
        form_type = 'hd'

        # Đường dẫn tới thư mục và file Excel
        data_dir = 'DataCollected'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        file_path = os.path.join(data_dir, 'data.xlsx')

        # Mở hoặc tạo mới file Excel
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            workbook.save(file_path)

        # Mapping headers to column indexes
        headers = {}
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, cell in enumerate(header_row):
            if cell:
                headers[cell.strip().lower()] = idx + 1  # 1-based index

        required_headers = ["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5","hd.01", "hd.02", "hd.03","pb"]
        # Kiểm tra các header cần thiết
        missing_headers = [h for h in required_headers if h not in headers]
        if missing_headers:
            # Nếu thiếu header, thêm lại header chuẩn
            sheet.delete_rows(1)
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            
            workbook.save(file_path)
            headers = {cell.strip().lower(): idx +1 for idx, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))}

        # Duyệt qua từng sinh viên và lưu thông tin vào file Excel
        for student in students:
            msv = student['msv']
            # Tìm sinh viên theo msv ở cột 'msv'
            student_found = False
            for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
                cell_msv = row[headers['msv'] -1]
                if cell_msv.value == msv:
                    student_found = True
                    # Tìm cột hd.01 đến hd.03 trống
                    for hd_col in ["hd.01", "hd.02", "hd.03"]:
                        cell_hd = row[headers[hd_col] -1]
                        if not cell_hd.value:
                            cell_hd.value = f"{lecturer_name} - C2.3: {student['diemC23']} - C3.2: {student['diemC32']} - C4.1: {student['diemC41']} - C6.1: {student['diemC61']} - C6.2: {student['diemC62']} - GPA: {student['gpa']}"
                            break  # Ghi xong vào cột trống, dừng tìm cột tiếp theo
                    break  # Đã tìm thấy sinh viên, dừng tìm kiếm

            if not student_found:
                # Thêm sinh viên mới và ghi vào hd.01
                new_row = [
                    student['fullname'],
                    student['msv'],
                    student['class'],
                    "",
                    "",
                    "",
                    "",
                    "",
                    f"{lecturer_name} - C2.3: {student['diemC23']} - C3.2: {student['diemC32']} - C4.1: {student['diemC41']} - C6.1: {student['diemC61']} - C6.2: {student['diemC62']} - GPA: {student['gpa']}",
                    "",
                    "",
                    ""
                ]
                sheet.append(new_row)
                print(f"Thêm sinh viên mới: {msv} và ghi dữ liệu vào hd.01")

        # Lưu file Excel
        workbook.save(file_path)

        # Chuyển hướng đến trang testOutput.html với dữ liệu
        return render(request, 'testOutput.html', {
            'students': students,
            'nhanXet': nhanXet,
            'lecturer_name': lecturer_name,
            'project_type': project_type,
            'project_name': project_name,
            'form_type': form_type
        })

    return redirect('huongdan3')


def process_form_hdcm(request):
    if request.method == 'POST':
        # Lấy số lượng sinh viên
        try:
            students_count = int(request.POST.get('students_count', 0))
        except ValueError:
            students_count = 0

        # Lấy dữ liệu từ form
        students = []
        for i in range(1, students_count + 1):
            student = {
                'fullname': request.POST.get(f'student_fullname_{i}', '').strip(),
                'msv': request.POST.get(f'student_msv_{i}', '').strip(),
                'class': request.POST.get(f'student_class_{i}', '').strip(),
                'diemC33': request.POST.get(f'diemC33SV{i}', '').strip(),
                'diemC42': request.POST.get(f'diemC42SV{i}', '').strip(),
                'diemC53': request.POST.get(f'diemC53SV{i}', '').strip(),
                'diemC63': request.POST.get(f'diemC63SV{i}', '').strip(),
                'diemC64': request.POST.get(f'diemC64SV{i}', '').strip(),
                'gpa': request.POST.get(f'gpaSV{i}', '').strip()
            }
            if student['msv']:  # Đảm bảo có mã sinh viên
                students.append(student)

        nhanXet = request.POST.get('nhanXet', '').strip()
        lecturer_name = request.POST.get('lecturer_name', '').strip()
        project_type = request.POST.get('project_type', '').strip()
        project_name = request.POST.get('project_name', '').strip()
        form_type = 'hd'

        # Đường dẫn tới thư mục và file Excel
        data_dir = 'DataCollected'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        file_path = os.path.join(data_dir, 'data.xlsx')

        # Mở hoặc tạo mới file Excel
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            workbook.save(file_path)

        # Mapping headers to column indexes
        headers = {}
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, cell in enumerate(header_row):
            if cell:
                headers[cell.strip().lower()] = idx + 1  # 1-based index

        required_headers = ["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5","hd.01", "hd.02", "hd.03","pb"]
        # Kiểm tra các header cần thiết
        missing_headers = [h for h in required_headers if h not in headers]
        if missing_headers:
            # Nếu thiếu header, thêm lại header chuẩn
            sheet.delete_rows(1)
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            
            workbook.save(file_path)
            headers = {cell.strip().lower(): idx +1 for idx, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))}

        # Duyệt qua từng sinh viên và lưu thông tin vào file Excel
        for student in students:
            msv = student['msv']
            # Tìm sinh viên theo msv ở cột 'msv'
            student_found = False
            for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
                cell_msv = row[headers['msv'] -1]
                if cell_msv.value == msv:
                    student_found = True
                    # Tìm cột hd.01 đến hd.03 trống
                    for hd_col in ["hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5"]:
                        cell_hd = row[headers[hd_col] -1]
                        if not cell_hd.value:
                            cell_hd.value = f"{lecturer_name} - C3.3: {student['diemC33']} - C4.2: {student['diemC42']} - C5.3: {student['diemC53']} - C6.3: {student['diemC63']} - C6.4: {student['diemC64']} - GPA: {student['gpa']}"
                            break  # Ghi xong vào cột trống, dừng tìm cột tiếp theo
                    break  # Đã tìm thấy sinh viên, dừng tìm kiếm

            if not student_found:
                # Thêm sinh viên mới và ghi vào hd.01
                new_row = [
                    student['fullname'],
                    student['msv'],
                    student['class'],
                    f"{lecturer_name} - C3.3: {student['diemC33']} - C4.2: {student['diemC42']} - C5.3: {student['diemC53']} - C6.3: {student['diemC63']} - C6.4: {student['diemC64']} - GPA: {student['gpa']}",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    ""
                ]
                sheet.append(new_row)
                # print(f"Thêm sinh viên mới: {msv} và ghi dữ liệu vào hd.01")

        # Lưu file Excel
        workbook.save(file_path)

        # Chuyển hướng đến trang testOutput.html với dữ liệu
        return render(request, 'testOutput.html', {
            'students': students,
            'nhanXet': nhanXet,
            'lecturer_name': lecturer_name,
            'project_type': project_type,
            'project_name': project_name,
            'form_type': form_type
        })

    return redirect('hoiDongChuyenMon')

def process_form_pb(request):
    if request.method == 'POST':
        # Lấy số lượng sinh viên
        try:
            students_count = int(request.POST.get('students_count', 0))
        except ValueError:
            students_count = 0

        # Lấy dữ liệu từ form
        students = []
        for i in range(1, students_count + 1):
            student = {
                'fullname': request.POST.get(f'student_fullname_{i}', '').strip(),
                'msv': request.POST.get(f'student_msv_{i}', '').strip(),
                'class': request.POST.get(f'student_class_{i}', '').strip(),
                'diemC23': request.POST.get(f'diemC23SV{i}', '').strip(),
                'diemC32': request.POST.get(f'diemC32SV{i}', '').strip(),
                'diemC41': request.POST.get(f'diemC41SV{i}', '').strip(),
                'diemC61': request.POST.get(f'diemC61SV{i}', '').strip(),
                'diemC62': request.POST.get(f'diemC62SV{i}', '').strip(),
                'gpa': request.POST.get(f'gpaSV{i}', '').strip()
            }
            if student['msv']:  # Đảm bảo có mã sinh viên
                students.append(student)

        nhanXet = request.POST.get('nhanXet', '').strip()
        lecturer_name = request.POST.get('lecturer_name', '').strip()
        project_type = request.POST.get('project_type', '').strip()
        project_name = request.POST.get('project_name', '').strip()
        form_type = 'hd'

        # Đường dẫn tới thư mục và file Excel
        data_dir = 'DataCollected'
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)

        file_path = os.path.join(data_dir, 'data.xlsx')

        # Mở hoặc tạo mới file Excel
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Sheet1"
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            workbook.save(file_path)

        # Mapping headers to column indexes
        headers = {}
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, cell in enumerate(header_row):
            if cell:
                headers[cell.strip().lower()] = idx + 1  # 1-based index

        required_headers = ["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv5","hd.01", "hd.02", "hd.03","pb"]
        # Kiểm tra các header cần thiết
        missing_headers = [h for h in required_headers if h not in headers]
        if missing_headers:
            # Nếu thiếu header, thêm lại header chuẩn
            sheet.delete_rows(1)
            sheet.append(["name", "msv", "class", "hdcm.uv1","hdcm.uv2","hdcm.uv3","hdcm.uv4","hdcm.uv55","hd.01", "hd.02", "hd.03","pb"])
            
            workbook.save(file_path)
            headers = {cell.strip().lower(): idx +1 for idx, cell in enumerate(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True)))}

        # Duyệt qua từng sinh viên và lưu thông tin vào file Excel
        for student in students:
            msv = student['msv']
            # Tìm sinh viên theo msv ở cột 'msv'
            student_found = False
            for row in sheet.iter_rows(min_row=2, max_col=12, values_only=False):
                cell_msv = row[headers['msv'] -1]
                if cell_msv.value == msv:
                    student_found = True
                    # Tìm cột hd.01 đến hd.03 trống
                    for hd_col in ["pb"]:
                        cell_hd = row[headers[hd_col] -1]
                        if not cell_hd.value:
                            cell_hd.value = f"{lecturer_name} - C2.3: {student['diemC23']} - C3.2: {student['diemC32']} - C4.1: {student['diemC41']} - C6.1: {student['diemC61']} - C6.2: {student['diemC62']} - GPA: {student['gpa']}"
                            break  # Ghi xong vào cột trống, dừng tìm cột tiếp theo
                    break  # Đã tìm thấy sinh viên, dừng tìm kiếm

            if not student_found:
                # Thêm sinh viên mới và ghi vào hd.01
                new_row = [
                    student['fullname'],
                    student['msv'],
                    student['class'],
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    f"{lecturer_name} - C2.3: {student['diemC23']} - C3.2: {student['diemC32']} - C4.1: {student['diemC41']} - C6.1: {student['diemC61']} - C6.2: {student['diemC62']} - GPA: {student['gpa']}"
                ]
                sheet.append(new_row)
                # print(f"Thêm sinh viên mới: {msv} và ghi dữ liệu vào hd.01")

        # Lưu file Excel
        workbook.save(file_path)

        # Chuyển hướng đến trang testOutput.html với dữ liệu
        return render(request, 'testOutput.html', {
            'students': students,
            'nhanXet': nhanXet,
            'lecturer_name': lecturer_name,
            'project_type': project_type,
            'project_name': project_name,
            'form_type': form_type
        })

    return redirect('canBoPhanBien')
