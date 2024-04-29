from django.shortcuts import render
from django.http import HttpResponse
import openpyxl
import os


def index(request):
    if request.method == 'POST':
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # Construct the path to the Excel file
        #excel_file_path = os.path.join(base_dir, 'data', 'plan.xls')

        # Load the Excel file
        #wb = openpyxl.load_workbook(excel_file_path)
        #sheet = wb['DT']
        #total_no_tests = sheet.max_row - 1
        total_no_tests = int(request.POST.get('total_no_tests', 0))
        duration_per_test = int(request.POST.get('Duration_per_test', 0))
        no_of_slots = int(request.POST.get('no_of_slots', 0))
        no_of_hrs_per_day = int(request.POST.get('no_of_hrs_per_day', 0))
        qual_cycle = int(request.POST.get('qual_cycle', 0))
        no_of_days = qual_cycle * 7
        optimized_days = ((total_no_tests * duration_per_test) / ((no_of_hrs_per_day * 60) * no_of_slots))
        optimized_duts = ((total_no_tests * duration_per_test) / ((qual_cycle * 7) * (no_of_hrs_per_day * 60)))

        # Return the result to the user
        return render(request, 'result.html', {'optimized_days': optimized_days, 'no_of_slots': no_of_slots,
                                               'optimized_duts': optimized_duts, 'no_of_days': no_of_days})
    else:
        # If it's a GET request, render the form
        return render(request, 'index.html')
