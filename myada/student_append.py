# Import necessary libraries
import pandas as pd
from django.http import HttpResponse
from django.shortcuts import render, get_object_or_404
from myada.models import Student, StudentResult # Update with your actual app and model names

# Function to fetch data from the Django database
def get_data_from_database(student):
    results = StudentResult.objects.filter(student=student)
    
    results_values = results.values()
    for result in results_values:
        result['created_at'] = result['created_at'].strftime('%Y-%m-%d %H:%M:%S')
        result['updated_at'] = result['updated_at'].strftime('%Y-%m-%d %H:%M:%S')
    
    data = pd.DataFrame.from_records(results.values())
    
    print("Data before conversion:")
    print(data)
    
    return data

# Function to append data to an existing Excel template starting from the third row
# Function to append data to an existing Excel template starting from the third row
# and save the new file with a different name in the current directory
def append_data_to_excel(existing_excel_path, data, new_excel_filename):
    existing_data = pd.read_excel(existing_excel_path)

    # Determine the starting row for appending data (3rd row in this case)
    start_row = existing_data.shape[0] + 2  # 0-based index, so the third row is index 2

    # Append data starting from the determined row
    combined_data = pd.concat([existing_data, data], ignore_index=True)
    
    # Set the new file path
    new_excel_path = new_excel_filename 
    
    # Save the new file with a different name in the current directory
    combined_data.to_excel(new_excel_path, index=False, startrow=start_row)


def student_view_results(request):
    student = get_object_or_404(Student, admin=request.user)

    # Fetch data from the database
    database_data = get_data_from_database(student)

    # Define paths for existing and new Excel files
    existing_excel_path = "test.xlsx"  # Assuming the file is in the current directory
    new_excel_filename = "new_test.xlsx"  # Desired name for the new Excel file

    # Append data to the new Excel file in the current directory
    append_data_to_excel(existing_excel_path, database_data, new_excel_filename)

    # Construct the full path for the new Excel file in the current directory
    new_excel_path = new_excel_filename

    # Serve the new Excel file for download
    response = HttpResponse(open(new_excel_path, 'rb').read())
    response['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response['Content-Disposition'] = f'attachment; filename={new_excel_filename}'

    return response