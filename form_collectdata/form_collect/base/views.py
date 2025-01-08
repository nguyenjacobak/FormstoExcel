from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook,load_workbook
import os
# Create your views here.
def index(request):
    if request.method == 'POST':
        name = request.POST.get('name')
        email = request.POST.get('email')

        # Define the directory and file path
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        data_dir = os.path.join(base_dir, 'DataCollected')
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        file_path = os.path.join(data_dir, 'data.xlsx')

        # Create or load the workbook
        if os.path.exists(file_path):
            workbook = load_workbook(file_path)
        else:
            workbook = Workbook()
            workbook.active.append(['Name', 'Email'])  # Add headers

        sheet = workbook.active
        sheet.append([name, email])

        workbook.save(file_path)
        return HttpResponse("Form submitted successfully!")

    return render(request, 'index.html')