from flask import Flask, request, jsonify
import gspread
import pandas as pd
import pyzipper
import tempfile
import os
import re
import base64

app = Flask(__name__)

# GoogleスプレッドシートのURLからIDを抽出する関数
def extract_id_from_url(url):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

# ZIPファイルを作成するエンドポイント
@app.route("/make_zip", methods=["POST"])
def make_zip_from_spreadsheet():
    data = request.get_json()
    url = data.get("file_url")  # スプレッドシートのURL
    password = data.get("password")  # ZIPファイルのパスワード

    # URLまたはパスワードが不足している場合
    if not url or not password:
        return jsonify({"error": "Missing file_url or password"}), 400

    # スプレッドシートIDの抽出
    spreadsheet_id = extract_id_from_url(url)
    if not spreadsheet_id:
        return jsonify({"error": "Invalid spreadsheet URL"}), 400

    try:
        # Google Sheets APIの認証
        gc = gspread.service_account(filename="credentials.json")  # サービスアカウントの認証
        sh = gc.open_by_key(spreadsheet_id)  # スプレッドシートを開く
        worksheet = sh.sheet1  # 最初のシートを選択
        df = pd.DataFrame(worksheet.get_all_records())  # スプレッドシートのデータをDataFrameに変換

        # 一時ディレクトリにファイルを保存
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = os.path.join(tmpdir, "result.xlsx")  # Excelファイルのパス
            zip_path = os.path.join(tmpdir, "result.zip")  # ZIPファイルのパス

            # Excelファイルの保存
            df.to_excel(excel_path, index=False)

            # ZIPファイルの作成（パスワード付き）
            with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                zf.setpassword(password.encode())  # パスワード設定
                zf.write(excel_path, arcname="result.xlsx")  # ExcelをZIPに追加

            # ZIPファイルをbase64エンコードして返す
            with open(zip_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")

            return jsonify({"zip_base64": encoded})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)  # 0.0.0.0でリッスンして外部からアクセス可能にする
