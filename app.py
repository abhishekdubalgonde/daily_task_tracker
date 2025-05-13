from flask import Flask, render_template, request, send_file
import os
import json
import pandas as pd
from datetime import datetime

app = Flask(__name__)
DATA_FOLDER = 'daily_logs'
REMARKS_FILE = 'data/remarks.json'

# Load remarks data
with open(REMARKS_FILE, 'r') as file:
    sub_remarks = json.load(file)

# Get today's filename
def get_today_filename():
    today = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(DATA_FOLDER, f"{today}.xlsx")

# Get next complaint ID
def get_next_id():
    file_path = get_today_filename()
    if not os.path.exists(file_path):
        return 1
    df = pd.read_excel(file_path)
    return len(df) + 1

@app.route('/', methods=['GET', 'POST'])
def index():
    categories = ["Software", "Hardware", "Network", "Internet"]
    sub_categories = [item["Sub Category"] for item in sub_remarks]

    if request.method == 'POST':
        today = datetime.now().strftime("%d/%b/%y")
        short_month = datetime.now().strftime("%b")
        count = get_next_id()

        complaint_id = f"SR/{short_month}/{count:03d}"
        start = request.form['start_time']
        end = request.form['end_time']
        effort_time = calculate_effort(start, end)

        sub_cat = request.form['sub_category']
        remark = next((item["Remarks"] for item in sub_remarks if item["Sub Category"] == sub_cat), "")

        row = {
            "Request/Complaint ID": complaint_id,
            "Created Date": today,
            "Start Time": start,
            "End Time": end,
            "User Name": request.form['user_name'],
            "Process": request.form['process'],
            "Reported By": request.form['reported_by'],
            "Priority": "Medium",
            "Technician Name": "Abhishek",
            "Issue Category": request.form['issue_category'],
            "Sub Category": sub_cat,
            "Effort Time": effort_time,
            "Request Status": "CLOSED",
            "Remarks": remark
        }

        file_path = get_today_filename()
        df_new = pd.DataFrame([row])

        if os.path.exists(file_path):
            df_existing = pd.read_excel(file_path)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new

        # Add Sl No
        if "Sl No" not in df_combined.columns:
            df_combined.insert(0, "Sl No", range(1, len(df_combined) + 1))
        else:
            df_combined["Sl No"] = range(1, len(df_combined) + 1)


        # Define exact column order
        column_order = [
            "Sl No", "Request/Complaint ID", "Created Date", "Start Time", "End Time",
            "User Name", "Process", "Reported By", "Priority", "Technician Name",
            "Issue Category", "Sub Category", "Effort Time", "Request Status", "Remarks"
        ]

        df_combined = df_combined[column_order]
        df_combined.to_excel(file_path, index=False)

    return render_template('index.html', categories=categories, sub_categories=sub_categories)

@app.route('/download')
def download():
    file_path = get_today_filename()
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "No file available for today."

def calculate_effort(start, end):
    fmt = "%H:%M"
    try:
        start_time = datetime.strptime(start, fmt)
        end_time = datetime.strptime(end, fmt)
        duration = end_time - start_time
        return f"{duration.seconds // 3600:02}:{(duration.seconds // 60) % 60:02}"
    except:
        return "00:00"

if __name__ == '__main__':
    os.makedirs(DATA_FOLDER, exist_ok=True)
    app.run(debug=True)
