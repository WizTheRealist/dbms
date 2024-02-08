from django.shortcuts import get_object_or_404
from django.http import HttpResponse
from openpyxl import Workbook
from .models import StudentResult, Student
import os

# views.py
from django.shortcuts import get_object_or_404
from django.http import HttpResponse
from openpyxl import Workbook
from .models import Student, StudentResult

def create_excel_sheets(request):
    # Get all students
    students = Student.objects.all()

    # Directory to store the generated spreadsheet
    storage_directory = 'student_spreadsheets'
    os.makedirs(storage_directory, exist_ok=True)  # Create directory if it doesn't exist

    for student in students:
        # Get all results for the current student
        student_results = StudentResult.objects.filter(student=student)

        wb = Workbook()  # Create a new Workbook (Excel file)
        ws = wb.active  # Select the active sheet

        # Assuming StudentResult model has fields like 'test' and 'exam'
        ws.append(['Course', 'Test', 'Exam', 'Total'])

        for student_result in student_results:
            
            total = student_result.test + student_result.exam
            # Retrieve information for the current student result
            result_info = [student_result.course.name, student_result.test, student_result.exam, total]

            # Add the student's result information to the spreadsheet
            ws.append(result_info)

        # Save the spreadsheet with the student's ID as part of the file name
        file_name = f"{student.admin.first_name}_{student.admin.last_name}_result.xlsx"
        file_path = os.path.join(storage_directory, file_name)

        # Save the workbook (Excel file) to the specified directory
        wb.save(file_path)

    return HttpResponse(f"Spreadsheet generated successfully.")

def view_spreadsheet(request):
    student = get_object_or_404(Student, admin=request.user)

    # Directory to store the generated spreadsheets
    storage_directory = 'student_spreadsheets'
    file_name = f"{student.admin.first_name}_{student.admin.last_name}_result.xlsx"
    file_path = os.path.join(storage_directory, file_name)

    if os.path.exists(file_path):
        with open(file_path, 'rb') as file:
            response = HttpResponse(file.read(), content_type='application/ms-excel')
            response['Content-Disposition'] = f'inline; filename="{file_name}"'
            return response
    else:
        return HttpResponse(f"Spreadsheet for {student.admin.last_name}, {student.admin.first_name} not found.")
