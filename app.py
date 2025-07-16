from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
from pymongo import MongoClient, DESCENDING
from pymongo.errors import ConnectionFailure, ConfigurationError
from datetime import datetime, time
import os
import uuid

# --- ตั้งค่าพื้นฐาน ---
app = Flask(__name__)
# สร้างโฟลเดอร์สำหรับเก็บไฟล์ชั่วคราว
# ใช้ path ที่สัมพันธ์กับตำแหน่งของไฟล์ app.py เพื่อความแน่นอน
basedir = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(basedir, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- ตั้งค่าการเชื่อมต่อ MongoDB ---
# แก้ไข Connection URI ของคุณได้ที่นี่
MONGO_URI = "mongodb://nbtcuser:Channipa@192.168.0.11:27017/myDatabase?authSource=admin"


# --- ฟังก์ชันประมวลผลข้อมูล (แยกออกมาเพื่อใช้ซ้ำ) ---
def process_excel_data(filepath, selected_date_str):
    """
    อ่านไฟล์ Excel, กรองข้อมูล, และสร้าง list ของ records ที่จะนำเข้า
    """
    # 1. อ่านไฟล์ Excel
    df = pd.read_excel(filepath, sheet_name='IssueTracker', header=None, skiprows=8)
    if df.empty:
        raise ValueError("ไม่พบข้อมูลในชีท IssueTracker (ตั้งแต่แถวที่ 9 เป็นต้นไป)")

    # 2. กรองข้อมูลตามวันที่เลือก
    selected_date = datetime.strptime(selected_date_str, '%Y-%m-%d')
    date_column_index = 6 # inform_date ถูกเลื่อนไปที่คอลัมน์ G (index 6)
    df[date_column_index] = pd.to_datetime(df[date_column_index], errors='coerce').dt.date
    filtered_df = df[df[date_column_index] == selected_date.date()].copy()

    if filtered_df.empty:
        raise ValueError(f"ไม่พบข้อมูลสำหรับวันที่ {selected_date.strftime('%d/%m/%Y')} ในชีท IssueTracker")

    # 3. สร้าง log_id
    client = None
    try:
        # เพิ่ม Timeout และทดสอบการเชื่อมต่อ
        client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
        client.admin.command('ping') # ทดสอบการเชื่อมต่อกับเซิร์ฟเวอร์

        db = client['nbtc']
        collection = db['service_request_nbtc']
        last_doc = collection.find_one(
            {"assignment_date": selected_date.strftime('%Y-%m-%d')},
            sort=[("log_id", DESCENDING)]
        )
    except (ConnectionFailure, ConfigurationError) as e:
        # ดักจับ Error การเชื่อมต่อและแสดงข้อความที่เข้าใจง่าย
        raise ValueError(f"ไม่สามารถเชื่อมต่อฐานข้อมูลได้ กรุณาตรวจสอบ Connection URI (Error: {e})")
    finally:
        if client:
            client.close()

    last_sequence = 0
    if last_doc and 'log_id' in last_doc and last_doc['log_id']:
        try:
            last_sequence = int(last_doc['log_id'].split('-')[-1])
        except (ValueError, IndexError):
            last_sequence = 0

    # 4. สร้างรายการข้อมูล (documents)
    records_to_insert = []
    now = datetime.now()
    # สร้างรูปแบบวันที่สำหรับ log_id (ปี ค.ศ. 2 หลัก YYMMDD)
    date_for_log_id = selected_date.strftime('%y%m%d')

    def format_date_to_string(value):
        if pd.isna(value): return None
        if isinstance(value, (datetime, pd.Timestamp)): return value.strftime('%Y-%m-%d')
        return str(value)

    def format_time_to_string(value):
        if pd.isna(value): return None
        if isinstance(value, time): return value.strftime('%H:%M')
        if isinstance(value, (datetime, pd.Timestamp)): return value.strftime('%H:%M')
        return str(value)

    def calculate_actual_time(start_time_str, end_time_str):
        if not start_time_str or not end_time_str: return None
        try:
            FMT = '%H:%M'
            dummy_date = datetime.min
            start_dt = datetime.combine(dummy_date.date(), datetime.strptime(start_time_str, FMT).time())
            end_dt = datetime.combine(dummy_date.date(), datetime.strptime(end_time_str, FMT).time())
            if end_dt < start_dt: return None
            delta = end_dt - start_dt
            hours = int(delta.total_seconds() // 3600)
            minutes = int((delta.total_seconds() % 3600) // 60)
            return f"{hours}:{minutes}"
        except (ValueError, TypeError):
            return None

    for row_tuple in filtered_df.itertuples(index=False, name=None):
        row = list(row_tuple)
        last_sequence += 1

        def get_value(index, default=None):
            try:
                val = row[index]
                return None if pd.isna(val) else val
            except IndexError:
                return default

        assignment_time_str = format_time_to_string(get_value(7)) # Column H
        completed_time_str = format_time_to_string(get_value(9)) # Column J

        record = {
            "actual_time": calculate_actual_time(assignment_time_str, completed_time_str),
            "log_id": f"PJ-NBT009-SS-{date_for_log_id}-{last_sequence:03d}",
            "assignment_date": format_date_to_string(selected_date),
            "assignment_description": get_value(4), # Column E
            "assignment_time": assignment_time_str, # Column H
            "completed_date": format_date_to_string(pd.to_datetime(get_value(8), errors='coerce')), # Column I
            "completed_time": completed_time_str, # Column J
            "create_date": now,
            "create_user": "AutoImportExcel",
            "inform_date": format_date_to_string(pd.to_datetime(get_value(6), errors='coerce')), # Column G
            "inform_time": assignment_time_str, # Column H
            "informer": get_value(11), # Column L
            "informer_contact": get_value(11), # <-- แก้ไข: ดึงข้อมูลจาก Column L เหมือนกับ informer
            "informer_department": None,
            "informer_email": None,
            "issue_details": get_value(2), # Column C
            "issue_img_1": None, "issue_img_2": None, "issue_img_3": None,
            "issue_type": "operation",
            "main_issue": "ลูกค้าแจ้งขอการสนับสนุนการดำเนินงาน",
            "operator": None, "operator_contact": None,
            "project_code": "PJ-NBT009",
            "project_name": "Any Registration",
            "recipient": get_value(10), # Column K
            "recipient_contact": get_value(10), # Column K
            "service_id": None,
            "service_status": "Complete",
            "service_type": "SS",
            "sla": "48",
            "update_date": now,
            "update_user": "AutoImportExcel"
        }
        records_to_insert.append(record)

    return records_to_insert


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    """
    หน้าแรกสำหรับอัปโหลดไฟล์ เมื่ออัปโหลดแล้วจะไปยังหน้า Preview
    """
    if request.method == 'POST':
        if 'excel_file' not in request.files:
            return render_template('index.html', error="ไม่พบไฟล์ที่แนบมา")

        file = request.files['excel_file']
        selected_date_str = request.form.get('selected_date')

        if file.filename == '':
            return render_template('index.html', error="กรุณาเลือกไฟล์ที่ต้องการอัปโหลด")

        if not selected_date_str:
            return render_template('index.html', error="กรุณาเลือกวันที่ที่ต้องการตรวจสอบ")

        if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            return render_template('index.html', error="รูปแบบไฟล์ไม่ถูกต้อง กรุณาอัปโหลดไฟล์ .xlsx หรือ .xls เท่านั้น")

        # บันทึกไฟล์ลงในตำแหน่งชั่วคราว
        temp_filename = f"{uuid.uuid4()}_{file.filename}"
        temp_filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        file.save(temp_filepath)

        try:
            # ประมวลผลข้อมูลเพื่อสร้างหน้า Preview
            records_for_preview = process_excel_data(temp_filepath, selected_date_str)
            return render_template(
                'preview.html',
                records=records_for_preview,
                temp_filename=temp_filename,
                selected_date=selected_date_str
            )
        except Exception as e:
            # หากเกิด Error ให้ลบไฟล์ชั่วคราวและแสดงข้อความ
            if os.path.exists(temp_filepath):
                os.remove(temp_filepath)
            return render_template('index.html', error=f"เกิดข้อผิดพลาด: {e}")

    return render_template('index.html')


@app.route('/confirm', methods=['POST'])
def confirm_import():
    """
    รับข้อมูลจากหน้า Preview เพื่อยืนยันการนำเข้าข้อมูลลง Database
    """
    temp_filename = request.form.get('temp_filename')
    selected_date_str = request.form.get('selected_date')

    if not all([temp_filename, selected_date_str]):
        return render_template('index.html', error="Session หมดอายุหรือคำขอไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")

    temp_filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)

    if not os.path.exists(temp_filepath):
        return render_template('index.html', error="ไม่พบไฟล์ชั่วคราว กรุณาลองใหม่อีกครั้ง")

    client = None
    try:
        # ประมวลผลไฟล์อีกครั้งเพื่อความถูกต้องของข้อมูล
        records_to_insert = process_excel_data(temp_filepath, selected_date_str)

        # เชื่อมต่อและบันทึกข้อมูล
        if records_to_insert:
            client = MongoClient(MONGO_URI)
            db = client['nbtc']
            collection = db['service_request_nbtc']
            collection.insert_many(records_to_insert)

        return redirect(url_for('success_page', count=len(records_to_insert)))

    except Exception as e:
        return render_template('index.html', error=f"เกิดข้อผิดพลาดระหว่างการยืนยันข้อมูล: {e}")
    finally:
        # ปิดการเชื่อมต่อและลบไฟล์ชั่วคราวเสมอ
        if client:
            client.close()
        if os.path.exists(temp_filepath):
            os.remove(temp_filepath)


@app.route('/success')
def success_page():
    count = request.args.get('count', 0, type=int)
    return render_template('success.html', record_count=count)


if __name__ == '__main__':
    # การเพิ่ม host='0.0.0.0' จะทำให้คอมพิวเตอร์เครื่องอื่นในเครือข่ายเดียวกัน
    # สามารถเข้าถึงเว็บแอปพลิเคชันนี้ได้ผ่าน IP Address ของเครื่องที่รันอยู่
    app.run(host='0.0.0.0', debug=True)
