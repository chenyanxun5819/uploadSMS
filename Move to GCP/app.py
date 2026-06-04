from flask import Flask, request, send_file, render_template_string
import subprocess
import os

app = Flask(__name__)

TEMPLATE_PATH = 'template.xlsx'
UPLOAD_PATH = 'uploaded.xlsx'

HTML = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>學術成績自動上傳</title>
</head>
<body>
    <h2>學術成績自動上傳</h2>
    <form method="POST" enctype="multipart/form-data">
        <label>下載範例 Excel：</label>
        <a href="/download-template">下載 template.xlsx</a><br><br>
        <label>上傳填寫好的 Excel：</label>
        <input type="file" name="file" accept=".xlsx" required><br><br>
        <button type="submit">開始自動上傳</button>
    </form>
    {% if result %}
    <p>{{ result }}</p>
    {% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    if request.method == 'POST':
        f = request.files['file']
        f.save(UPLOAD_PATH)
        # 執行 upload.py
        try:
            output = subprocess.check_output(['python', 'upload.py', UPLOAD_PATH], stderr=subprocess.STDOUT, timeout=60)
            result = output.decode('utf-8')
        except Exception as e:
            result = str(e)
    return render_template_string(HTML, result=result)

@app.route('/download-template')
def download_template():
    return send_file(TEMPLATE_PATH, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
