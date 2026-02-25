"""
app.py - Sales Kit AI Agent 主程式
Flask 後端，處理表單提交、呼叫生成器、提供下載
"""

import os
import uuid
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from generator import generate_sales_kit
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 最大上傳 16MB
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.secret_key = os.environ.get('SECRET_KEY', 'sales-kit-dev-secret')

# 確保必要資料夾存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 允許的圖片格式
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def allowed_file(filename):
    """檢查檔案副檔名是否被允許"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """首頁：顯示 Sales Kit 生成表單"""
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    """接收表單，呼叫 AI 生成 PPTX，回傳下載頁面"""
    try:
        # 取得必填表單欄位
        kit_type = request.form.get('kit_type', 'product')
        company_name = request.form.get('company_name', '').strip()
        core_value = request.form.get('core_value', '').strip()
        target_audience = request.form.get('target_audience', '').strip()
        data_highlights = request.form.get('data_highlights', '').strip()
        contact_info = request.form.get('contact_info', '').strip()
        extra_info = request.form.get('extra_info', '').strip()

        # 基本驗證
        if not company_name or not core_value:
            raise ValueError("請填寫公司名稱與核心訴求")

        # 處理 Logo 上傳
        logo_path = None
        if 'logo' in request.files:
            file = request.files['logo']
            if file and file.filename and allowed_file(file.filename):
                ext = file.filename.rsplit('.', 1)[1].lower()
                safe_name = f"logo_{uuid.uuid4().hex[:8]}.{ext}"
                logo_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
                file.save(logo_path)

        # 組裝傳給生成器的資料
        form_data = {
            'kit_type': kit_type,
            'company_name': company_name,
            'core_value': core_value,
            'target_audience': target_audience,
            'data_highlights': data_highlights,
            'contact_info': contact_info,
            'extra_info': extra_info,
            'logo_path': logo_path,
        }

        # 生成輸出檔名
        safe_company = secure_filename(company_name[:20]) or 'saleskit'
        output_filename = f"{safe_company}_{uuid.uuid4().hex[:6]}.pptx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # 呼叫 AI 生成器（核心步驟）
        generate_sales_kit(form_data, output_path)

        # 回傳成功結果頁面
        kit_type_names = {
            'bni': 'BNI 會員招募',
            'product': '產品／服務銷售',
            'brand': '品牌提案',
            'event': '活動推廣',
        }
        return render_template(
            'result.html',
            success=True,
            filename=output_filename,
            company_name=company_name,
            kit_type_label=kit_type_names.get(kit_type, kit_type),
        )

    except Exception as e:
        # 發生錯誤：回傳友善錯誤頁面
        return render_template(
            'result.html',
            success=False,
            error=str(e),
        )


@app.route('/download/<filename>')
def download(filename):
    """提供 PPTX 檔案下載"""
    # 安全性：只允許從 output 資料夾下載，防止路徑遍歷攻擊
    safe_name = secure_filename(filename)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], safe_name)

    if not os.path.exists(output_path):
        return render_template('result.html', success=False, error="檔案不存在或已過期，請重新生成"), 404

    return send_file(
        output_path,
        as_attachment=True,
        download_name=safe_name,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('FLASK_DEBUG', 'false').lower() == 'true')
