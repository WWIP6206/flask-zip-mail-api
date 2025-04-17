from flask import Flask, request, jsonify
import gspread
import pandas as pd
import pyzipper
import tempfile
import os
import re
import base64

app = Flask(__name__)

def extract_id_from_url(url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

@app.route("/make_zip", methods=["POST"])
def make_zip_from_spreadsheet():
    try:
        # ✅ request.get_json() に戻す！
        data = request.get_json()
    except Exception as e:
        return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400

    url = data.get("file_url")
    password = data.get("password")

    if not url or not password:
        return jsonify({"error": "Missing file_url or password"}), 400

    spreadsheet_id = extract_id_from_url(url)
    if not spreadsheet_id:
        return jsonify({"error": "Invalid spreadsheet URL"}), 400

    try:
        gc = gspread.service_account(filename="/etc/secrets/credentials.json")
        sh = gc.open_by_key(spreadsheet_id)
        worksheet = sh.sheet1
        df = pd.DataFrame(worksheet.get_all_records())

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = os.path.join(tmpdir, "result.xlsx")
            zip_path = os.path.join(tmpdir, "result.zip")

            df.to_excel(excel_path, index=False)

            with pyzipper.AESZipFile(zip_path, 'w',
                                     compression=pyzipper.ZIP_DEFLATED,
                                     encryption=pyzipper.WZ_AES) as zf:
                zf.setpassword(password.encode())
                zf.write(excel_path, arcname="result.xlsx")

            with open(zip_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")

        return jsonify({"zip_base64": encoded})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

