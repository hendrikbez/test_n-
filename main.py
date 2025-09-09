from flask import Flask, request, jsonify, session, redirect, render_template, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import os
import win32com.client
import pythoncom
from pathlib import Path
import time
import shutil
import threading
import json
from datetime import datetime, timedelta
import locale

app = Flask(__name__)
app.secret_key = 'my-private-register-app-2025'

# Excel file path
EXCEL_FILE_PATH = r"C:\AutoHotkey\Kerk\ninja\2025_Register_Naamlys.xlsm"
EXCEL_BACKUP_PATH = r"C:\AutoHotkey\Kerk\ninja\2025_Register_Naamlys_backup.xlsm"

if not os.path.exists(EXCEL_FILE_PATH):
    print(f"WARNING: Excel file not found at C:\AutoHotkey\Kerk\ninja\2025_Register_Naamlys.xlsm")
else:
    print(f"Excel file found at C:\AutoHotkey\Kerk\ninja\2025_Register_Naamlys.xlsm")



# User credentials
USERS = {
    'admin': generate_password_hash('6371')
   
}

# Data storage
DATA_FILE = 'church_register.json'



def get_afrikaans_date():
    # Dictionary for Afrikaans month names
    afrikaans_months = {
        1: "Januarie", 2: "Februarie", 3: "Maart", 4: "April", 
        5: "Mei", 6: "Junie", 7: "Julie", 8: "Augustus", 
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    
    today = datetime.now()
    day = today.day
    month = afrikaans_months[today.month]
    year = today.year
    
    # Format as "05 September 2025"
    return f"{day:02d} {month} {year}"

# Function to update the Voorblad date
def update_voorblad_date():
    print("Starting update_voorblad_date()")
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Excel
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Open the workbook
            wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
            
            # Get the Voorblad sheet (Sheet4)
            try:
                sheet = wb.Worksheets('Voorblad')
            except:
                sheet = wb.Worksheets('Sheet4')  # Fallback if sheet name is different
            
            # Get formatted date
            afrikaans_date = get_afrikaans_date()
            
            # Update cells G14:N16 with the date
            # Merge the cells first if they aren't already merged
            try:
                sheet.Range("G14:N16").MergeCells = True
            except:
                pass  # Already merged
                
            # Set the value and formatting
            sheet.Range("G14:N16").Value = afrikaans_date
            
            # Updated formatting:
            sheet.Range("G14:N16").HorizontalAlignment = -4131  # xlLeft (-4131)
            sheet.Range("G14:N16").VerticalAlignment = -4108  # xlCenter
            sheet.Range("G14:N16").Font.Name = "Cambria"
            sheet.Range("G14:N16").Font.Size = 28
            sheet.Range("G14:N16").Font.Bold = True
            
            # Save the workbook
            wb.Save()
            wb.Close()
            excel.Quit()
            
            # Clean up COM
            pythoncom.CoUninitialize()
            
            return True
        except Exception as e:
            print(f"Error updating Voorblad date: {e}")
            if excel:
                try:
                    wb.Close(False)  # Close without saving
                    excel.Quit()
                except:
                    pass
            return False
    except Exception as e:
        print(f"COM initialization error: {e}")
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def excel_date_to_string(excel_date):
    """Convert Excel date number to string in dd-mmm format"""
    try:
        # Excel dates are number of days since 1900-01-01
        dt = datetime(1900, 1, 1) + timedelta(days=excel_date - 2)
        return dt.strftime('%d-%b')
    except:
        return str(excel_date)

def get_data_via_com():
    """Get data using Excel COM interface to avoid VBA corruption"""
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        sheet = wb.Worksheets('Register')
        
        data = []
        current_family_id = None
        
        # Find the last row with data
        last_row = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row  # -4162 = xlUp
        
        # Start from row 11 instead of row 2
        start_row = 11
        
        for row in range(start_row, last_row + 1):
            # Skip rows that are completely empty or are headers
            if sheet.Cells(row, 2).Value is None:
                continue
                
            # Get family ID (column A)
            family_id = sheet.Cells(row, 1).Value
            if family_id is not None and family_id != "":
                current_family_id = family_id
                
            # Get number (column B)
            number = sheet.Cells(row, 2).Value
            if number is not None:
                # Format dates properly
                verj = sheet.Cells(row, 5).Value
                huwelik = sheet.Cells(row, 6).Value
                
                # Convert Excel dates to proper format
                if verj and isinstance(verj, float):
                    try:
                        verj = excel_date_to_string(verj)
                    except:
                        verj = str(verj)
                
                if huwelik and isinstance(huwelik, float):
                    try:
                        huwelik = excel_date_to_string(huwelik)
                    except:
                        huwelik = str(huwelik)
                
                data.append({
                    'familyId': current_family_id,
                    'number': number,
                    'van': sheet.Cells(row, 3).Value or '',
                    'naam': sheet.Cells(row, 4).Value or '',
                    'verj': verj or '',
                    'huwelik': huwelik or '',
                    'selfoon': sheet.Cells(row, 7).Value or '',
                    'adres': sheet.Cells(row, 8).Value or '',
                    'epos': sheet.Cells(row, 9).Value or ''  # Added email column
                })
        
        wb.Close(False)  # Close without saving
        excel.Quit()
        return jsonify({'data': data})
    except Exception as e:
        if excel:
            try:
                excel.Quit()
            except:
                pass
        pythoncom.CoUninitialize()
        raise e

def get_data_via_openpyxl():
    """Fallback to openpyxl if COM fails"""
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH, keep_vba=True, data_only=True)
    sheet = wb['Register']
    
    data = []
    current_family_id = None
    
    # Start from row 11
    start_row = 11
    
    for row_idx, row in enumerate(sheet.iter_rows(min_row=start_row, values_only=True), start=start_row):
        # Skip completely empty rows or header rows
        if all(cell is None for cell in row) or (row[1] is not None and isinstance(row[1], str) and row[1].upper() in ["NAAM", "NOMER", "NOMMER", "NUMBER"]):
            continue
            
        # Get family ID (column A) - use previous if empty
        family_id = row[0]
        if family_id is not None and family_id != "":
            current_family_id = family_id
            
        # Only add rows with data in column B (number)
        if row[1] is not None and not (isinstance(row[1], str) and row[1].upper() in ["NAAM", "NOMER", "NOMMER", "NUMBER"]):
            # Format dates properly
            verj = row[4]
            huwelik = row[5]
            
            # Convert Excel dates to proper format
            if verj and isinstance(verj, datetime):
                verj = verj.strftime('%d-%b')
            elif verj and isinstance(verj, str) and len(verj) > 10:
                # Try to parse date strings
                try:
                    dt = datetime.strptime(verj[:10], '%Y-%m-%d')
                    verj = dt.strftime('%d-%b')
                except:
                    pass
            
            if huwelik and isinstance(huwelik, datetime):
                huwelik = huwelik.strftime('%d-%b')
            elif huwelik and isinstance(huwelik, str) and len(huwelik) > 10:
                # Try to parse date strings
                try:
                    dt = datetime.strptime(huwelik[:10], '%Y-%m-%d')
                    huwelik = dt.strftime('%d-%b')
                except:
                    pass
            
            data.append({
                'familyId': current_family_id,
                'number': row[1],
                'van': row[2] or '',
                'naam': row[3] or '',
                'verj': verj or '',
                'huwelik': huwelik or '',
                'selfoon': row[6] or '',
                'adres': row[7] or '',
                'epos': row[8] or '' if len(row) > 8 else ''  # Added email column
            })
    
    wb.close()
    return jsonify({'data': data})

def add_data_operation(sheet, data):
    """Operation to add data via COM"""
    try:
        last_row = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
        next_row = last_row + 1
        
        # Write data to columns C-I only, leaving A-B for macros
        sheet.Cells(next_row, 3).Value = data['van']
        sheet.Cells(next_row, 4).Value = data['naam']
        sheet.Cells(next_row, 5).Value = data['verj']
        sheet.Cells(next_row, 6).Value = data['huwelik']
        sheet.Cells(next_row, 7).Value = data['selfoon']
        sheet.Cells(next_row, 8).Value = data['adres']
        sheet.Cells(next_row, 9).Value = data.get('epos', '')  # Added email column
        
        # Return success but don't set a number since we're not setting column B
        return True, None
    except Exception as e:
        print(f"Add operation error: {e}")
        return False, None

def add_family_operation(sheet, data):
    """Operation to add family via COM"""
    try:
        last_row = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
        next_row = last_row + 1
        
        members = data['members']
        family_adres = data.get('familyAdres', '')
        
        # Write each family member (only columns C-I)
        for i, member in enumerate(members):
            row_num = next_row + i
            
            # Don't set family ID or number (columns A-B)
            sheet.Cells(row_num, 3).Value = member['van']
            sheet.Cells(row_num, 4).Value = member['naam']
            sheet.Cells(row_num, 5).Value = member['verj']
            
            # Only include wedding date for the first member
            if i == 0:
                sheet.Cells(row_num, 6).Value = member.get('huwelik', '')
            else:
                sheet.Cells(row_num, 6).Value = ""
            
            sheet.Cells(row_num, 7).Value = member['selfoon']
            sheet.Cells(row_num, 8).Value = family_adres
            sheet.Cells(row_num, 9).Value = member.get('epos', '')  # Added email column
        
        # Return success but don't set a family ID since we're not setting column A
        return True, None
    except Exception as e:
        print(f"Add family operation error: {e}")
        return False, None

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        'register': [],  # Main member data
        'next_id': 1
    }

def save_data(data):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@app.route('/')
def index():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    return render_template('login.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username in USERS and check_password_hash(USERS[username], password):
            session['user'] = username
            
            # Update the Voorblad date when user logs in
            try:
                update_voorblad_date()
            except Exception as e:
                print(f"Failed to update Voorblad date: {e}")
                # Continue even if date update fails
                
            return jsonify({'success': True})
        return jsonify({'success': False, 'message': 'Ongeldige aanmelding'})
    
    return render_template('login.html')


@app.route('/add_person', methods=['POST'])
def add_person():
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    
    try:
        data = request.json
        
        if not data.get('van') or not data.get('naam'):
            return jsonify({'success': False, 'message': 'Van en naam is verpligtend'}), 400
        
        # Format phone number if provided
        selfoon = data.get('selfoon', '')
        if selfoon:
            digits = ''.join(filter(str.isdigit, selfoon))
            if len(digits) >= 10:
                selfoon = f"{digits[:3]} {digits[3:6]} {digits[6:10]}"

        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
            sheet = wb.Worksheets('Register')
            
            # --- CORRECTED LOGIC TO FIND NEXT ROW ---
            # Find the last row based on data in Column C ('Van'), which is more reliable
            last_row = sheet.Cells(sheet.Rows.Count, "C").End(-4162).Row
            next_row = last_row + 1
            
            # Write data to the new row
            sheet.Cells(next_row, 3).Value = data['van']
            sheet.Cells(next_row, 4).Value = data['naam']
            sheet.Cells(next_row, 5).Value = data.get('verj', '')
            sheet.Cells(next_row, 6).Value = data.get('huwelik', '')
            sheet.Cells(next_row, 7).Value = selfoon
            sheet.Cells(next_row, 8).Value = data.get('adres', '')
            sheet.Cells(next_row, 9).Value = data.get('epos', '')
            
            wb.Save()
            wb.Close()
            excel.Quit()
            pythoncom.CoUninitialize()
            
            return jsonify({'success': True, 'message': 'Persoon suksesvol bygevoeg'})
        except Exception as e:
            print(f"Error adding person: {e}")
            if excel:
                try:
                    wb.Close(False)
                    excel.Quit()
                except: pass
            return jsonify({'success': False, 'message': str(e)})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'success': False, 'message': str(e)})
    finally:
        try:
            pythoncom.CoUninitialize()
        except: pass


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

@app.route('/dashboard')
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('dashboard.html')



@app.route('/run_macro', methods=['POST'])
def run_macro():
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    
    data = request.json
    macro_name = data.get('macroName')
    
    if not macro_name:
        return jsonify({'success': False, 'message': 'Geen makro naam verskaf nie'}), 400

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        # --- CRITICAL FIX: Disable events to prevent interference ---
        excel.EnableEvents = False
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        
        # Define the full macro path
        # The VBA code shows the macros are in the ThisWorkbook module, not a specific sheet module.
        full_macro_name = f"'{os.path.basename(EXCEL_FILE_PATH)}'!{macro_name}"
        print(f"Attempting to run macro: {full_macro_name}")

        # Run the macro
        excel.Application.Run(full_macro_name)
        
        message = f"Makro '{macro_name}' suksesvol uitgevoer."
        
        # The macro should handle saving, but we save again to be sure.
        wb.Save()
        wb.Close()
        
        # --- CRITICAL FIX: Re-enable events before quitting ---
        excel.EnableEvents = True
        excel.Quit()
        pythoncom.CoUninitialize()
        
        return jsonify({'success': True, 'message': message})
    
    except Exception as e:
        print(f"Error running macro '{macro_name}': {e}")
        if excel:
            # Ensure Excel is closed properly on error
            excel.EnableEvents = True
            try:
                wb.Close(False) # Close without saving changes
            except: pass
            excel.Quit()
        return jsonify({'success': False, 'message': f"Fout met makro '{macro_name}': {e}"})
    
    finally:
        try:
            pythoncom.CoUninitialize()
        except: pass

@app.route('/get_current_date')
def get_current_date():
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Excel
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Open the workbook
            wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
            
            # Get the Voorblad sheet (Sheet4)
            try:
                sheet = wb.Worksheets('Voorblad')
            except:
                sheet = wb.Worksheets('Sheet4')  # Fallback if sheet name is different
            
            # Get the current date from the sheet
            current_date = sheet.Range("G14").Value
            
            wb.Close(False)  # Close without saving
            excel.Quit()
            
            # Clean up COM
            pythoncom.CoUninitialize()
            
            return jsonify({'success': True, 'date': current_date or get_afrikaans_date()})
        except Exception as e:
            print(f"Error getting date: {e}")
            if excel:
                try:
                    wb.Close(False)  # Close without saving
                    excel.Quit()
                except:
                    pass
            return jsonify({'success': False, 'message': str(e)})
    except Exception as e:
        print(f"COM initialization error: {e}")
        return jsonify({'success': False, 'message': str(e)})
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route('/update_date', methods=['POST'])
def update_date():
    print("update_date endpoint called")
    if 'user' not in session:
        print("User not authenticated")
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    
    try:
        print(f"Excel file path: {EXCEL_FILE_PATH}")
        print("Calling update_voorblad_date()")
        # Update the date in the Excel file
        success = update_voorblad_date()
        print(f"update_voorblad_date() returned: {success}")
        
        if success:
            current_date = get_afrikaans_date()
            print(f"Date updated successfully to: {current_date}")
            return jsonify({'success': True, 'date': current_date, 'message': 'Datum suksesvol opgedateer'})
        else:
            print("Failed to update date")
            return jsonify({'success': False, 'message': 'Kon nie datum opdateer nie'})
    except Exception as e:
        print(f"Error updating date: {e}")
        return jsonify({'success': False, 'message': str(e)})

@app.route('/get_data')
def get_data():
    if 'user' not in session:
        return jsonify({'error': 'Not authenticated'}), 401
    
    try:
        # First try to use COM for reading to preserve VBA integrity
        try:
            return get_data_via_com()
        except Exception as com_error:
            print(f"COM read failed, trying openpyxl: {com_error}")
            return get_data_via_openpyxl()
    except Exception as e:
        print(f"Error loading data: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/SortFamilyMembers', methods=['POST'])
def sort_family_members():
    if 'user' not in session:
        return jsonify({'success': False, 'message': 'Not authenticated'}), 401
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Excel
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Open the workbook
            wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
            
            # Get the Register sheet
            sheet = wb.Worksheets('Register')
            
            # Find the last row with data
            last_row = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
            
            # Define the range to sort (from row 11 to last_row, columns A to I)
            sort_range = sheet.Range(f"A11:I{last_row}")
            
            # Sort by Van (column C) and then by Naam (column D)
            sort_range.Sort(
                Key1=sheet.Range("C11"),
                Order1=1,  # xlAscending
                Key2=sheet.Range("D11"),
                Order2=1,  # xlAscending
                Header=1  # xlYes (include headers)
            )
            
            # Save the workbook
            wb.Save()
            wb.Close()
            excel.Quit()
            
            # Clean up COM
            pythoncom.CoUninitialize()
            
            return jsonify({'success': True, 'message': 'Register gesorteer volgens familienaam'})
        except Exception as e:
            print(f"Error sorting register: {e}")
            if excel:
                try:
                    wb.Close(False)  # Close without saving
                    excel.Quit()
                except:
                    pass
            return jsonify({'success': False, 'message': str(e)})
    except Exception as e:
        print(f"COM initialization error: {e}")
        return jsonify({'success': False, 'message': str(e)})
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass



@app.route('/some_long_operation')
def long_operation():
    try:
        # Do some work
        session['progress'] = 25
        # More work
        session['progress'] = 50
        # Final work
        session['progress'] = 100
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})






 # Add this to your main.py after the imports
@app.before_request
def before_request():
    if 'progress' not in session:
        session['progress'] = 0

@app.after_request
def after_request(response):
    # Reset progress after each request
    session['progress'] = 0
    return response

@app.route('/api/progress')
def get_progress():
    return jsonify({'progress': session.get('progress', 0)})





@app.route('/register')
def register_view():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('register.html')

if __name__ == '__main__':
    print("Starting Flask server on http://127.0.0.1:5000")
    app.run(debug=True, port=5000)