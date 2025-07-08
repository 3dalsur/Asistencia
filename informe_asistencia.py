import pandas as pd
from datetime import datetime, timedelta
import locale
from fpdf import FPDF
from flask import Flask, request, render_template, redirect, url_for, send_from_directory
import os # Import the os module for path operations and file removal

app = Flask(__name__)

# Define a directory to save uploaded files and generated reports
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER) # Create the directory if it doesn't exist

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to read and process the file
def read_file(file_path):
    data = []
    with open(file_path, 'r', encoding='utf-8') as file:
        # Skip the header
        next(file)
        for line in file:
            row = line.strip().split('\t')
            # Ensure enough columns exist to prevent IndexError
            if len(row) > 6:
                # --- CHANGE THIS LINE ---
                # Original: datetime.strptime(record['DateTime'], '%Y-%m-%d %H:%M:%S')
                # Corrected:
                try:
                    # Attempt to parse with '/' first
                    dt_obj = datetime.strptime(row[6].strip(), '%Y/%m/%d %H:%M:%S')
                except ValueError:
                    # If that fails, try with '-' (for flexibility)
                    dt_obj = datetime.strptime(row[6].strip(), '%Y-%m-%d %H:%M:%S')

                data.append({
                    'EnNo': row[2].strip(),
                    'Name': row[3].strip(),
                    'DateTime': dt_obj.strftime('%Y-%m-%d %H:%M:%S') # Store in consistent format
                })
            else:
                print(f"Skipping malformed row: {line.strip()}")
    return data

# Function to generate report for each employee
def generate_report(data, selected_month):
    reports = {}
    # Set locale to Spanish
    # On Windows, 'es_ES' might not work directly. Try 'Spanish_Spain.1252' or '' (empty string for default)
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
        except locale.Error:
            locale.setlocale(locale.LC_TIME, '') # Fallback to default locale


    # Get the month name in Spanish
    date = datetime.strptime(selected_month, '%Y-%m')
    month_name = date.strftime('%B %Y').upper()

    # Filter data by selected month
    filtered_data = [record for record in data if
                     datetime.strptime(record['DateTime'], '%Y-%m-%d %H:%M:%S').strftime('%Y-%m') == selected_month]

    # Group by employee
    employees = {}
    for record in filtered_data:
        if record['EnNo'] not in employees:
            employees[record['EnNo']] = {'Name': record['Name'], 'records': []}
        employees[record['EnNo']]['records'].append(record)

    # Generate report for each employee
    for EnNo, info in employees.items():
        pdf = FPDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, f"Asistencia: {info['Name']} | Registro: {EnNo} | {month_name}", 0, 1, 'C')

        pdf.set_font('Arial', 'B', 12)
        pdf.cell(90, 10, "Mañana", 0, 0, 'C') # Adjusted width for "Mañana" header
        pdf.cell(90, 10, "Tarde", 0, 1, 'C')  # Adjusted width for "Tarde" header

        pdf.set_font('Arial', '', 10)
        # Headers for the first half of the table
        pdf.cell(15, 10, "Dia", 1, 0, 'C')
        pdf.cell(20, 10, "Semana", 1, 0, 'C')
        pdf.cell(25, 10, "Entrada", 1, 0, 'C')
        pdf.cell(25, 10, "Salida", 1, 0, 'C')
        pdf.cell(25, 10, "Entrada", 1, 0, 'C')
        pdf.cell(25, 10, "Salida", 1, 0, 'C')

        # Headers for the second half of the table
        pdf.cell(15, 10, "Dia", 1, 0, 'C')
        pdf.cell(20, 10, "Semana", 1, 0, 'C')
        pdf.cell(25, 10, "Entrada", 1, 0, 'C')
        pdf.cell(25, 10, "Salida", 1, 0, 'C')
        pdf.cell(25, 10, "Entrada", 1, 0, 'C')
        pdf.cell(25, 10, "Salida", 1, 1, 'C') # 1 at the end for new line

        records_by_day = {}
        for record in info['records']:
            try:
                date = datetime.strptime(record['DateTime'], '%Y-%m-%d %H:%M:%S')
                day = date.strftime('%d')
                weekday = date.strftime('%A')
                weekday_es = {
                    'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
                    'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
                }.get(weekday, weekday) # Use .get with a default for safety
                time = date.strftime('%H:%M:%S')

                if day not in records_by_day:
                    records_by_day[day] = {'weekday': weekday_es, 'times': []}
                # Add a check to prevent adding duplicate times if they are too close (within 30 seconds)
                if len(records_by_day[day]['times']) == 0 or (
                        datetime.strptime(time, '%H:%M:%S') - datetime.strptime(records_by_day[day]['times'][-1],
                                                                                '%H:%M:%S')).seconds > 30:
                    records_by_day[day]['times'].append(time)
            except ValueError:
                print(f"Skipping malformed DateTime in record: {record['DateTime']}")


        # Populate the table for days 1-31 (or until the end of the month)
        # Determine the number of days in the selected month
        num_days_in_month = (date.replace(month=date.month%12+1, day=1) - timedelta(days=1)).day
        
        for i in range(1, (num_days_in_month // 2) + (num_days_in_month % 2) + 1): # Adjust loop for correct rows
            day_str1 = str(i).zfill(2)
            day_str2 = str(i + num_days_in_month // 2).zfill(2) # Calculate the corresponding day in the second column

            # First column of the report
            if day_str1 in records_by_day:
                times1 = records_by_day[day_str1]['times']
                morning_entry1 = times1[0] if len(times1) > 0 else '-----'
                morning_exit1 = times1[1] if len(times1) > 1 else '-----'
                afternoon_entry1 = times1[2] if len(times1) > 2 else '-----'
                afternoon_exit1 = times1[3] if len(times1) > 3 else '-----'

                pdf.cell(15, 10, day_str1, 1, 0, 'C')
                pdf.cell(20, 10, records_by_day[day_str1]['weekday'], 1, 0, 'C')
                pdf.cell(25, 10, morning_entry1, 1, 0, 'C')
                pdf.cell(25, 10, morning_exit1, 1, 0, 'C')
                pdf.cell(25, 10, afternoon_entry1, 1, 0, 'C')
                pdf.cell(25, 10, afternoon_exit1, 1, 0, 'C')
            else:
                try:
                    current_date1 = datetime.strptime(f"{selected_month}-{day_str1}", '%Y-%m-%d')
                    weekday_es1 = {
                        'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
                        'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
                    }.get(current_date1.strftime('%A'), current_date1.strftime('%A'))
                except ValueError:
                    weekday_es1 = '-----' # Handle cases where day_str1 might be invalid for the month
                pdf.cell(15, 10, day_str1, 1, 0, 'C')
                pdf.cell(20, 10, weekday_es1, 1, 0, 'C')
                pdf.cell(25, 10, "-----", 1, 0, 'C')
                pdf.cell(25, 10, "-----", 1, 0, 'C')
                pdf.cell(25, 10, "-----", 1, 0, 'C')
                pdf.cell(25, 10, "-----", 1, 0, 'C')

            # Second column of the report (for days 16-31 or remaining days)
            if i + num_days_in_month // 2 <= num_days_in_month: # Only try to get if the day exists
                if day_str2 in records_by_day:
                    times2 = records_by_day[day_str2]['times']
                    morning_entry2 = times2[0] if len(times2) > 0 else '-----'
                    morning_exit2 = times2[1] if len(times2) > 1 else '-----'
                    afternoon_entry2 = times2[2] if len(times2) > 2 else '-----'
                    afternoon_exit2 = times2[3] if len(times2) > 3 else '-----'

                    pdf.cell(15, 10, day_str2, 1, 0, 'C')
                    pdf.cell(20, 10, records_by_day[day_str2]['weekday'], 1, 0, 'C')
                    pdf.cell(25, 10, morning_entry2, 1, 0, 'C')
                    pdf.cell(25, 10, morning_exit2, 1, 0, 'C')
                    pdf.cell(25, 10, afternoon_entry2, 1, 0, 'C')
                    pdf.cell(25, 10, afternoon_exit2, 1, 1, 'C') # 1 at the end for new line
                else:
                    try:
                        current_date2 = datetime.strptime(f"{selected_month}-{day_str2}", '%Y-%m-%d')
                        weekday_es2 = {
                            'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
                            'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
                        }.get(current_date2.strftime('%A'), current_date2.strftime('%A'))
                    except ValueError:
                        weekday_es2 = '-----'
                    pdf.cell(15, 10, day_str2, 1, 0, 'C')
                    pdf.cell(20, 10, weekday_es2, 1, 0, 'C')
                    pdf.cell(25, 10, "-----", 1, 0, 'C')
                    pdf.cell(25, 10, "-----", 1, 0, 'C')
                    pdf.cell(25, 10, "-----", 1, 0, 'C')
                    pdf.cell(25, 10, "-----", 1, 1, 'C') # 1 at the end for new line
            else: # If no second column day exists, just move to the next line
                pdf.ln()

        # Save the PDF to the UPLOAD_FOLDER
        report_filename = f"{EnNo}_{month_name.replace(' ', '_')}.pdf"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        pdf.output(report_path)
        reports[EnNo] = report_filename # Store only the filename, not full path

    return reports

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part"
        file = request.files['file']
        if file.filename == '':
            return "No selected file"

        selected_month = request.form['month']

        if file:
            # Save the uploaded file temporarily
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)

            try:
                data = read_file(file_path)
                reports = generate_report(data, selected_month)

                # Store reports in session or pass them as query parameters
                # For simplicity, we'll redirect and the user can check the 'uploads' folder
                # In a real app, you might want to show a list of downloadable reports
                return render_template('index.html', reports=reports) # Pass reports to the template
            except Exception as e:
                return f"Error processing file: {e}"
            finally:
                # Optionally remove the uploaded file after processing
                if os.path.exists(file_path):
                    os.remove(file_path)
    return render_template('index.html')

# New route to serve generated PDF files
@app.route('/reports/<filename>')
def download_report(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)