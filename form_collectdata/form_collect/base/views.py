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

    return redirect('baoCaoTienDoL3')


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