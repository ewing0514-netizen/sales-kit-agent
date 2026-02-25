"""
app.py - Sales Kit AI Agent 主程式
Flask 後端，處理表單提交、呼叫生成器、直接回傳 PPTX 下載
Vercel Serverless 相容版：PPTX 生成至記憶體 (BytesIO)，不依賴持久化磁碟
"""

import os
import uuid
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from generator import generate_sales_kit
from dotenv import load_dotenv

# 載入環境變數（本機開發用）
load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 最大 16MB
app.secret_key = os.environ.get('SECRET_KEY', 'sales-kit-dev-secret')

# Vercel 環境使用 /tmp，本機使用相對路徑
UPLOAD_DIR = '/tmp/uploads' if os.environ.get('VERCEL') else 'static/uploads'
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 允許的圖片格式
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}


def allowed_file(filename: str) -> bool:
    """檢查檔案副檔名是否被允許"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """首頁表單"""
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    """
    接收表單 → 呼叫 AI 生成 → 直接回傳 PPTX 二進位流
    前端使用 fetch 接收 Blob 並觸發瀏覽器下載
    Vercel Serverless 相容：不依賴跨請求的磁碟檔案
    """
    try:
        # 取得表單欄位
        kit_type     = request.form.get('kit_type', 'product')
        company_name = request.form.get('company_name', '').strip()
        core_value   = request.form.get('core_value', '').strip()
        target       = request.form.get('target_audience', '').strip()
        data_hl      = request.form.get('data_highlights', '').strip()
        contact      = request.form.get('contact_info', '').strip()
        extra        = request.form.get('extra_info', '').strip()

        # 基本驗證
        if not company_name or not core_value:
            return jsonify({'error': '請填寫公司名稱與核心訴求'}), 400

        # 處理 Logo 上傳（Vercel 用 /tmp，本機用 static/uploads）
        logo_path = None
        if 'logo' in request.files:
            f = request.files['logo']
            if f and f.filename and allowed_file(f.filename):
                ext = f.filename.rsplit('.', 1)[1].lower()
                fname = f'logo_{uuid.uuid4().hex[:8]}.{ext}'
                logo_path = os.path.join(UPLOAD_DIR, fname)
                f.save(logo_path)

        form_data = {
            'kit_type':        kit_type,
            'company_name':    company_name,
            'core_value':      core_value,
            'target_audience': target,
            'data_highlights': data_hl,
            'contact_info':    contact,
            'extra_info':      extra,
            'logo_path':       logo_path,
        }

        # 生成 PPTX 至記憶體 BytesIO，不寫入磁碟
        buffer = BytesIO()
        generate_sales_kit(form_data, buffer)
        buffer.seek(0)

        # 組裝下載檔名
        safe_co = secure_filename(company_name[:20]) or 'saleskit'
        filename = f"{safe_co}_saleskit.pptx"

        # 直接回傳 PPTX 二進位流，前端 fetch 取得後觸發下載
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        )

    except Exception as e:
        # 回傳 JSON 錯誤，前端 JS 顯示友善訊息
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port,
            debug=os.environ.get('FLASK_DEBUG', 'false').lower() == 'true')
