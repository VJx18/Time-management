from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile
import uuid
import shutil


import openpyxl
import os
import pandas as pd
import warnings
from openpyxl import load_workbook
import random


class HolidayTool:
    
    def __init__(self, path):
        self.path = path
        self.emp_list = []
        self.meta_data = None
        self.max_emp = 30
        
        self.bl_mapping = {
            2020: {
                'Baden-Württemberg': 4, 'Bayern (katholisch)': 5, 'Bayern': 6,
                'Berlin': 7, 'Brandenburg': 8, 'Bremen': 9, 'Hamburg': 10,
                'Hessen': 11, 'Mecklenburg-Vorpommern': 12, 'Niedersachsen': 13,
                'Nordrhein-Westfahlen': 14, 'Rheinland-Pfalz': 15, 'Saarland': 16,
                'Sachsen': 17, 'Sachsen-Anhalt': 18, 'Schleswig-Holstein': 19,
                'Thüringen': 20
            },
            2021: {
                'Baden-Württemberg': 25, 'Bayern (katholisch)': 26, 'Bayern': 27,
                'Berlin': 28, 'Brandenburg': 29, 'Bremen': 30, 'Hamburg': 31,
                'Hessen': 32, 'Mecklenburg-Vorpommern': 33, 'Niedersachsen': 34,
                'Nordrhein-Westfahlen': 35, 'Rheinland-Pfalz': 36, 'Saarland': 37,
                'Sachsen': 38, 'Sachsen-Anhalt': 39, 'Schleswig-Holstein': 40,
                'Thüringen': 41
            },
            2022: {
                'Baden-Württemberg': 46, 'Bayern (katholisch)': 47, 'Bayern': 48,
                'Berlin': 49, 'Brandenburg': 50, 'Bremen': 51, 'Hamburg': 52,
                'Hessen': 53, 'Mecklenburg-Vorpommern': 54, 'Niedersachsen': 55,
                'Nordrhein-Westfahlen': 56, 'Rheinland-Pfalz': 57, 'Saarland': 58,
                'Sachsen': 59, 'Sachsen-Anhalt': 60, 'Schleswig-Holstein': 61,
                'Thüringen': 62
            },
             2023: {
                'Baden-Württemberg': 67, 'Bayern (katholisch)': 68, 'Bayern': 69,
                'Berlin': 70, 'Brandenburg': 71, 'Bremen': 72, 'Hamburg': 73,
                'Hessen': 74, 'Mecklenburg-Vorpommern': 75, 'Niedersachsen': 76,
                'Nordrhein-Westfahlen': 77, 'Rheinland-Pfalz': 78, 'Saarland': 79,
                'Sachsen': 80, 'Sachsen-Anhalt': 81, 'Schleswig-Holstein': 82,
                'Thüringen': 83
            }
        }
        
        self.base_row = 1
        self.yr_offset = self.calc_offsets()
        self.fei_first = 4
        self.fei_last = 377
        self.ist_first = 15
        self.ist_last = 390
        self.emp_start_row = 4
        
        warnings.filterwarnings("ignore", message="Conditional Formatting extension is not supported and will be removed")
    
    def calc_offsets(self):
        yrs = sorted(self.bl_mapping.keys())
        return {y: self.base_row + (i * 36) for i, y in enumerate(yrs)}
    
    def get_metadata(self):
        try:
            self.meta_data = pd.read_excel(self.path, sheet_name="MA Übersicht")
            return True
        except Exception as e:
            raise Exception(f"Metadata error: {e}")
    
    def find_employees(self):
        if self.meta_data is None:
             raise Exception("No metadata available")
        count = 0
        for i, row in self.meta_data.iterrows():
            if count >= self.max_emp:
          
                break

            if pd.isna(row.get("Vorname", None)) or pd.isna(row.get("Nachname", None)) or pd.isna(row.get("Bundesland", None)):
                continue

            start_dt = row.iloc[19]
            end_dt = row.iloc[20]

            emp_info = {
                'full_name': f"{row['Vorname']} {row['Nachname']}",
                'state': row["Bundesland"],
                'orig_idx': i,
                'emp_num': count + 1,
                'row_offset': count,
                'start': start_dt,
                'end': end_dt
            }

            self.emp_list.append(emp_info)
            count += 1

           
    
    def process_employee_holidays(self, emp_data, wb):
        sheet = wb['IST Stunden']

        try:
            fei_file_path = os.path.join(os.path.dirname(__file__), "Feiertage.xlsx")
            fei_wb = load_workbook(ei_file_path, data_only=True)
            fei_sheet = fei_wb.active
        except Exception as e:
            raise Exception(f"Can't load holidays file: {e}")
            return False

        print(f"\nDoing {emp_data['full_name']} in {emp_data['state']}...")

        emp_row_pos = self.emp_start_row + emp_data['row_offset']

        start_dt = pd.to_datetime(emp_data.get('start'))
        end_dt = pd.to_datetime(emp_data.get('end'))

        if pd.isna(start_dt) or pd.isna(end_dt):
            print(f"  Missing dates for {emp_data['full_name']}")
            return False

        print(f"  Working: {start_dt} to {end_dt}")

        total_marked = 0
        total_skipped = 0

        for yr, state_rows in self.bl_mapping.items():
            if emp_data['state'] not in state_rows:
                print(f"  Can't find {emp_data['state']} for {yr}")
                continue

            src_row = state_rows[emp_data['state']]
            holiday_data = []
            
            for col in range(self.fei_first, self.fei_last + 1):
                val = fei_sheet.cell(row=src_row, column=col).value
                holiday_data.append(val)

            processed_holidays = []
            for v in holiday_data:
                if v is not None and str(v).strip() != "":
                    processed_holidays.append('f')
                else:
                    processed_holidays.append(None)

            target_start = self.yr_offset[yr] + 3
            target_row_num = target_start + emp_row_pos - 4
            date_row = target_start - 2

            marked_this_yr = 0
            skipped_this_yr = 0

            for idx in range(len(processed_holidays)):
                if idx >= len(processed_holidays):
                    break
                    
                col_num = self.ist_first + idx
                if col_num > self.ist_last:
                    break

                if processed_holidays[idx] == 'f':
                    dt_cell = sheet.cell(row=date_row, column=col_num).value
                    dt_obj = pd.to_datetime(dt_cell, errors='coerce')

                    if pd.isna(dt_obj) or dt_obj < start_dt or dt_obj > end_dt:
                        skipped_this_yr += 1
                        continue

                    sheet.cell(row=target_row_num, column=col_num).value = "f"
                    marked_this_yr += 1

            print(f"  {yr}: marked {marked_this_yr}, skipped {skipped_this_yr}")
            total_marked += marked_this_yr
            total_skipped += skipped_this_yr

        print(f"  Total for {emp_data['full_name']}: {total_marked} marked, {total_skipped} skipped")
        return True
    
    def do_all_holidays(self, out_path=None):
        if not self.emp_list:
            print("No employees yet")
            return None
        
        print(f"\nProcessing {len(self.emp_list)} people...")
        
        try:
            wb = load_workbook(self.path, data_only=True)
        except Exception as e:
            print(f"Can't load file: {e}")
            return None
        
        good_count = 0
        for emp in self.emp_list:
            if self.process_employee_holidays(emp, wb):
                good_count += 1
            else:
                print(f"Failed on {emp['full_name']}")
        
        print(f"\nDone {good_count} of {len(self.emp_list)}")
        
        if out_path is None:
            out_path = self.path.replace('.xlsx', '_holidays_added.xlsx')
        
        try:
            wb.save(out_path)
            print(f"Saved to: {out_path}")
            return out_path
        except Exception as e:
            print(f"Save error: {e}")
            return None
    
    def execute(self, out_path=None):
        
        
        print("Getting metadata...")
        if not self.get_metadata():
            print("Metadata failed")
            return None
        
        print("\nFinding employees...")
        self.find_employees()
        
        if not self.emp_list:
            print("No employees found")
            return None
        
        print(f"\nGot {len(self.emp_list)} people:")
        for e in self.emp_list:
            print(f"  #{e['emp_num']}: {e['full_name']} ({e['state']})")
        
        print(f"\nMarking holidays...")
        result = self.do_all_holidays(out_path)
        
        if result:
            print(f"\nDone! Processed {len(self.emp_list)} employees")
            print(f"File: {result}")
        else:
            print("\nFailed")
        
        return result
    
    def change_max(self, new_max):
        if new_max < 1 or new_max > 50:
            print(f"Max should be 1-50, got {new_max}")
            return False
        
        self.max_emp = new_max
        print(f"Max set to: {new_max}")
        return True

app = Flask(__name__)
CORS(app)  
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 
temp_files = {}


def get_file_input():
    """Get file name from user with validation"""
    while True:
        print("\n=== File Selection ===")
        file_name = input("Enter Excel file name (or path): ").strip()
        
        if not file_name:
            print("Please enter a file name")
            continue
            
        # Add .xlsx if not present
        if not file_name.endswith('.xlsx'):
            file_name += '.xlsx'
            
        # Check if file exists
        if os.path.exists(file_name):
            print(f"Found file: {file_name}")
            return file_name
        else:
            print(f"File '{file_name}' not found")
            retry = input("Try again? (y/n): ").lower()
            if retry != 'y':
                return None




@app.route('/api/process-holidays', methods=['POST'])
def process_holidays():
    try:
        holiday_lines = [
            "Spreading holidays like cheese on a hot pizza 🍕📅",
            "Distributing holidays like Nutella on warm toast 🧈📆",
            "Layering holidays like frosting on a cake 🎂📅",
            "Smearing holidays across your calendar like butter on bread 🧈🗓️",
            "Sprinkling holidays like herbs on a Margherita 🍃🍕"
        ]
        fun_message = random.choice(holiday_lines)

        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not file.filename.endswith('.xlsx'):
            return jsonify({'error': 'Currently only accepting Excel Files'}), 400
        
      
        temp_dir = tempfile.mkdtemp()
        input_file = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(input_file)
        
   
        max_employees = request.form.get('max_employees', 30, type=int)
        
     
        tool = HolidayTool(input_file)
        tool.change_max(max_employees)
        
       
        output_filename = f"processed_{uuid.uuid4().hex}.xlsx"
        output_file = os.path.join(temp_dir, output_filename)
        
        result = tool.execute(output_file)
        
        if result:
  
            file_id = str(uuid.uuid4())
            temp_files[file_id] = {
                'path': result,
                'filename': output_filename,
                'temp_dir': temp_dir
            }
            
            return jsonify({
                'success': True,
                'message': fun_message,
                'employees_processed': len(tool.emp_list),
                'download_url': f'/api/download/{file_id}'
            })
        else:
           
            shutil.rmtree(temp_dir, ignore_errors=True)
            return jsonify({'error': 'Processing failed'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<file_id>')
def download_file(file_id):
    if file_id not in temp_files:
        return jsonify({'error': 'File not found'}), 404
    
    file_info = temp_files[file_id]
    try:
        return send_file(
            file_info['path'], 
            as_attachment=True, 
            download_name=file_info['filename']
        )
    except Exception as e:
        return jsonify({'error': 'Download failed'}), 500

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
