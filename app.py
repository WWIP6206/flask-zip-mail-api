from flask import Flask, request, jsonify
import openpyxl
import yagmail
import pyzipper
import os
import tempfile

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "GET":
        return "Flask server is running!"  # 動作確認用

    elif request.method == "POST":
        # JSONリクエスト取得
        data = request.get_json()
        sheet_name = data.get("sheet_name", "Sheet")
        values = data.get("values", [])

        # Excel・ZIPを一時ファイルで生成
        with tempfile.TemporaryDirectory() as tempdir:
            excel_path = os.path.join(tempdir, "data.xlsx")
            zip_path = os.path.join(tempdir, "data.zip")

            # Excelファイル作成
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name
            for row in values:
                ws.append(row)
            wb.save(excel_path)

            # パスワード付きZIP作成
            password = os.urandom(8).hex()
            with pyzipper.AESZipFile(zip_path, 'w', compression=pyzipper.ZIP_DEFLATED,
                                     encryption=pyzipper.WZ_AES) as zf:
                zf.setpassword(password.encode())
                zf.write(excel_path, arcname="data.xlsx")

            # メール設定（Gmail使用）
            sender_email = "あなたのGmailアドレス"
            app_password = "アプリパスワード"
            recipient_email = "受信者のメールアドレス"

            yag = yagmail.SMTP(sender_email, app_password)

            # 1通目：ZIP送信
            yag.send(
                to=recipient_email,
                subject="データファイル（ZIP付き）",
                contents="ZIPファイルを送信します。\nパスワードは別途お知らせします。",
                attachments=zip_path
            )

            # 2通目：パスワード送信
            yag.send(
                to=recipient_email,
                subject="ZIPファイルのパスワード",
                contents=f"ZIPファイルのパスワードは：\n\n{password}"
            )

        return jsonify({"status": "success", "message": "ファイル送信完了"})

if __name__ == "__main__":
    app.run(port=5000)
