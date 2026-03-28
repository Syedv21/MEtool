from flask import Flask, render_template, request, jsonify, send_from_directory, session, redirect, url_for
import cv2
import numpy as np
import easyocr
import pandas as pd
import os
from datetime import datetime
import openpyxl
app = Flask(__name__)
app.secret_key = "techbridge_secure_key_2026"  # Key for session management

# Initialize OCR (Optimized for standard CPU)
reader = easyocr.Reader(['en'], gpu=False)

# --- CONFIGURATION ---
# Matches your G: drive screenshot exactly
DRIVE_ROOT = r"G:\My Drive\Mortgage_Production"

HEADERS = [
    "Record Number", "Lead ID", "Applicant First Name", "Applicant Last Name",
    "Street Address", "City", "Post Code", "Applicant DOB",
    "Co Applicant First Name", "Co Applicant Last Name", "Best Time to Call", "Remarks",
    "Type Of Property", "Property Value", "Purpose of Loan", "Mortgage Type",
    "Loan Amount", "Loan Term", "Interest Type", "Monthly Installment",
    "Existing Loan", "Annual Income", "Down Payment", "Remarks 2",
    "Lender Name", "Loan Officer First Name", "Loan Officer Last Name",
    "T.R #", "N.I #", "Occupation", "Other Income",
    "Credit Card Issuer", "Credit Card Type", "Credit Score", "Remarks 3"
]


# --- ROUTES ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        path = request.form.get('path').strip()  # e.g., User5/Maaz
        if path:
            session['agent_path'] = path
            return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('agent_path', None)
    return redirect(url_for('login'))


@app.route('/')
def index():
    if 'agent_path' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', agent=session['agent_path'])


@app.route('/history')
def history():
    return render_template('production.html')


@app.route('/upload', methods=['POST'])
def upload():
    try:
        file_bytes = np.frombuffer(request.files['file'].read(), np.uint8)
        img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        h, w, _ = img.shape

        # Split image into 2 columns for Zig-Zag flow
        left_col = img[:, 0:w // 2]
        right_col = img[:, w // 2:w]

        all_words = []
        for col in [left_col, right_col]:
            res = reader.readtext(col, detail=0)
            all_words.extend(res)

        records = []
        current_record = []
        end_markers = ["EXCELLENT", "GOOD", "FAIR", "POOR", "EXCELLANT"]

        for word in all_words:
            if word.upper().startswith("TBY"):
                if current_record: records.append(" | ".join(current_record))
                current_record = [word]
            else:
                current_record.append(word)
                if any(m in word.upper() for m in end_markers):
                    records.append(" | ".join(current_record))
                    current_record = []

        if current_record: records.append(" | ".join(current_record))
        return jsonify({"records": records})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/save', methods=['POST'])
def save():
    if 'agent_path' not in session:
        return jsonify({"error": "Unauthorized"}), 401
    try:
        payload = request.get_json()
        raw_text = payload.get('text', "")

        # Split and clean
        data = [v.strip() for v in raw_text.split("|")]

        # CRITICAL: Ensure we have exactly 35 columns.
        # If user provided 30, add 5 empties. If 40, cut at 35.
        if len(data) < 35:
            data.extend([""] * (35 - len(data)))
        else:
            data = data[:35]

        agent_path = session['agent_path']
        target_dir = os.path.join(DRIVE_ROOT, agent_path.replace("/", "\\"))
        os.makedirs(target_dir, exist_ok=True)

        # File naming by Date
        filename = os.path.join(target_dir, f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
        df = pd.DataFrame([data], columns=HEADERS)

        if os.path.exists(filename):
            try:
                # Use engine='openpyxl' to avoid errors with G: Drive file locks
                old_df = pd.read_excel(filename, engine='openpyxl')
                pd.concat([old_df, df], ignore_index=True).to_excel(filename, index=False, engine='openpyxl')
            except Exception as e:
                # If file is locked or corrupted, create a new one with a timestamp
                filename = os.path.join(target_dir, f"{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx")
                df.to_excel(filename, index=False, engine='openpyxl')
        else:
            df.to_excel(filename, index=False, engine='openpyxl')

        return jsonify({"status": "success"})
    except Exception as e:
        print(f"SAVE ERROR: {str(e)}") # Check your terminal for this!
        return jsonify({"error": str(e)}), 500


@app.route('/list_files')
def list_files():
    # 1. Check if the agent is logged in
    if 'agent_path' not in session:
        return jsonify({"files": []})

    all_files = []
    # Get the specific path for this agent (e.g., "User5/Maaz")
    agent_folder_name = session['agent_path']

    # Construct the full physical path on your G: Drive
    # Result: G:\My Drive\Mortgage_Production\User5\Maaz
    specific_path = os.path.join(DRIVE_ROOT, agent_folder_name.replace("/", "\\"))

    # 2. Only scan THIS agent's folder
    if os.path.exists(specific_path):
        for file in os.listdir(specific_path):
            if file.endswith(".xlsx"):
                all_files.append({
                    "folder": agent_folder_name,
                    "filename": file
                })

    # Sort so the newest date is at the top
    all_files.sort(key=lambda x: x['filename'], reverse=True)
    return jsonify({"files": all_files})


@app.route('/download/<path:filepath>')
def download(filepath):
    # The <path:> converter allows the URL to include slashes (e.g., User1/Jyothi/file.xlsx)
    return send_from_directory(DRIVE_ROOT, filepath, as_attachment=True)

@app.route('/delete/<path:filepath>', methods=['DELETE'])
def delete_file(filepath):
    # WARNING: This allows any user to delete any file.
    try:
        full_path = os.path.join(DRIVE_ROOT, filepath.replace("/", "\\"))
        if os.path.exists(full_path):
            os.remove(full_path)
            return jsonify({"status": "success"})
        return jsonify({"status": "error", "message": "File not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    # host 0.0.0.0 allows all 10 agents on your WiFi to connect
    app.run(debug=True, host='0.0.0.0', port=5000)