import os
import re
import json
from PIL import Image
from paddleocr import PaddleOCR
from pdf2image import convert_from_path
from openpyxl import Workbook
import openai
from flask import Flask, request, render_template, send_from_directory, jsonify

# 替换为你的 OpenAI API 密钥
openai.api_key = "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"

# 系统提示语
system_prompt = '''
我将给你一段话，你需要从中：
1. 商品类别。例如：乳制品，方便食品，熟食制品，软饮料，其他食品。通常会在**之间。
2. 发票号码。通常以24开头20位或12位。
3. 发票代码。只有发票号码是12位的情况才会有发票代码，若有发票以24为开头的发票号码则不填写内容。
4. 价税合计金额。通常在字符‘小写’附近。
5. 不含税金额。通常只比价税合计金额小一些，注意不要其他符号只要数字。
6. 开票日期。年月日用"-"连接，例如，2024-01-01
7. 发票类型。你只能在以下三个中选择填写。电子发票-普票，增值税普通发票，增值税专用发票。

注意！严格按照以下内容及格式回答！返回一个JSON格式的对象，例如：
{
    "商品类别": "其他食品",
    "发票号码": "24...",
    "发票代码": "",
    "价税合计金额": "100.00",
    "不含税金额": "80.00",
    "开票日期": "2024-01-01",
    "发票类型": "电子发票-普票"
}
'''

# 初始化PaddleOCR模型，支持中文识别
ocr = PaddleOCR(use_angle_cls=True, lang='ch')

# Flask 初始化
app = Flask(__name__)

def convert_to_pdf(file_path, output_dir):
    file_name, file_extension = os.path.splitext(file_path)
    if file_extension.lower() == '.pdf':
        return

    output_pdf = os.path.join(output_dir, os.path.basename(file_name) + '.pdf')
    try:
        if file_extension.lower() in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            image = Image.open(file_path)
            image.convert('RGB').save(output_pdf, "PDF")
            print(f"已将 {file_path} 转换为 {output_pdf}")
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")


def extract_text_from_image(image_path):
    result = ocr.ocr(image_path, cls=True)
    all_text = ''.join([line[1][0].strip() for line in result[0]])
    return all_text


def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path, first_page=1, last_page=1)
    for img in images:
        img_path = '../temp_image.jpg'
        img.save(img_path, 'JPEG')
        all_text = extract_text_from_image(img_path)
        return all_text


def chat_with_gpt(user_prompt):
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini-2024-07-18",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        response_format={
            "type": "json_object"
        }
    )
    answer = response['choices'][0]['message']['content']
    try:
        answer_json = json.loads(answer)
    except json.JSONDecodeError:
        print("返回的内容无法解析为JSON格式：", answer)
        answer_json = {}
    return answer_json


def extract_invoice_info_from_pdf(pdf_path):
    all_text = extract_text_from_pdf(pdf_path)
    print(f"提取的OCR文本: {all_text}")
    info = chat_with_gpt(all_text)
    print("提取的信息:", info)
    return info


def rename_pdf_file(pdf_path, invoice_code, invoice_number, amount, output_dir):
    invoice_code_part = f"{invoice_code}-" if invoice_code else ""
    pdf_filename = f"{invoice_code_part}{invoice_number}-{amount}.pdf"
    new_pdf_path = os.path.join(output_dir, pdf_filename)
    os.rename(pdf_path, os.path.join(output_dir, new_pdf_path))
    print(f"重命名成功：{new_pdf_path}")
    return new_pdf_path


def process_folder_and_export_to_excel(input_dir, output_excel):
    wb = Workbook()
    ws = wb.active
    ws.append(["商品类别", "发票号码", "发票代码", "价税合计金额", "不含税金额", "开票日期", "发票类型", "PDF文件名"])

    # 转换非PDF文件为PDF
    for filename in os.listdir(input_dir):
        file_path = os.path.join(input_dir, filename)
        if os.path.isfile(file_path) and not filename.lower().endswith('.pdf'):
            convert_to_pdf(file_path, input_dir)

    # 处理所有PDF文件
    for filename in os.listdir(input_dir):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(input_dir, filename)
            print(f"正在处理文件: {pdf_path}")
            info = extract_invoice_info_from_pdf(pdf_path)
            if info:
                product_category = info.get("商品类别", "")
                invoice_number = info.get("发票号码", "")
                invoice_code = info.get("发票代码", "")
                total_amount = info.get("价税合计金额", "")
                tax_excluded_amount = info.get("不含税金额", "")
                invoice_date = info.get("开票日期", "")
                invoice_type = info.get("发票类型", "")
                ws.append([product_category, invoice_number, invoice_code, total_amount, tax_excluded_amount, invoice_date, invoice_type, filename])
                wb.save(output_excel)
                print(f"已保存至Excel: {output_excel}")
                if invoice_number and total_amount:
                    rename_pdf_file(pdf_path, invoice_code, invoice_number, total_amount, output_folder)

    print(f"所有数据已写入Excel: {output_excel}")

# 定义Flask路由
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    if file:
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        return jsonify({"message": "File uploaded successfully", "filename": file.filename})

@app.route('/process', methods=['POST'])
def process():
    filename = request.form.get('filename')
    if not filename:
        return "No filename provided", 400
    pdf_path = os.path.join('uploads', filename)
    if not os.path.exists(pdf_path):
        return "File not found", 404
    info = extract_invoice_info_from_pdf(pdf_path)
    return jsonify(info)

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory('uploads', filename)

# 运行Flask应用
if __name__ == "__main__":
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True, host='0.0.0.0', port=5555)

# 在同一目录下创建templates文件夹，并在其中创建index.html文件。
# index.html应包含一个简单的文件上传表单和一个处理按钮。
