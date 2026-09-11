import os
import re
import pandas as pd
import pdfplumber
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "secret_key_for_invoice_parser"

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def parse_pdf(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # 1. Invoice No. (例如: Invoice no. : 10161)
    inv_no_match = re.search(
        r"Invoice\s*no\.?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_no = inv_no_match.group(1).strip() if inv_no_match else ""

    # 2. Invoice Date (例如: Invoice date: 09/08/2026)
    date_match = re.search(
        r"Invoice\s*date\s*[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 3. PO Number (例如: PO#: PSDELT260902005)
    po_match = re.search(
        r"PO\s*#?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    po_number = po_match.group(1).strip() if po_match else ""

    # 4. Total Amount (例如: Total $800.00)
    total_match = re.search(r"Total\s*[\$]?\s*([0-9,]+\.[0-9]{2})", full_text)
    total_amount = total_match.group(1).strip() if total_match else ""

    # 5. SKU (例如: SKU# OLSAC504AA400-001)
    sku_match = re.search(
        r"SKU\s*#?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    sku = sku_match.group(1).strip() if sku_match else ""

    # 6 & 7. FROM & TO（紅色螢光筆部分）
    # 對應文字格式: "... from 601 Delta Plano to 1991 Peak Smart on 9/3/2026 ..."
    from_val = ""
    to_val = ""

    # 第一種邏輯：容許中間有或沒有 Plano，直接抓取完整標記字串
    route_match = re.search(
        r"from\s+(.*?)\s+(?:Plano\s+)?to\s+(.*?)(?:\s+on|\s+Dock|\n|$)",
        full_text,
        re.IGNORECASE,
    )

    if route_match:
        from_val = route_match.group(1).strip()  # 結果為: 601 Delta
        to_val = route_match.group(2).strip()  # 結果為: 1991 Peak Smart
    else:
        # 備援規則：若沒有寫在一行內，嘗試分開比對
        from_alt = re.search(
            r"from\s+([0-9]+\s+[A-Za-z]+)", full_text, re.IGNORECASE
        )
        to_alt = re.search(
            r"to\s+([0-9]+\s+[A-Za-z\s]+?)(?:\s+on|\n|$)",
            full_text,
            re.IGNORECASE,
        )
        if from_alt:
            from_val = from_alt.group(1).strip()
        if to_alt:
            to_val = to_alt.group(1).strip()

    return {
        "Invoice No": invoice_no,
        "Invoice Date": invoice_date,
        "PO Number": po_number,
        "From": from_val,
        "To": to_val,
        "SKU": sku,
        "Total Amount": total_amount,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "files" not in request.files:
            flash("請選擇檔案")
            return redirect(request.url)

        files = request.files.getlist("files")
        if not files or files[0].filename == "":
            flash("未選擇任何檔案")
            return redirect(request.url)

        data_list = []
        for file in files:
            if file and file.filename.lower().endswith(".pdf"):
                filename = secure_filename(file.filename)
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                file.save(filepath)

                try:
                    res = parse_pdf(filepath)
                    res["Filename"] = filename
                    data_list.append(res)
                except Exception as e:
                    print(f"處理檔案 {filename} 時出錯: {e}")

        if not data_list:
            flash("未能成功解析上傳的 PDF 檔案")
            return redirect(request.url)

        # 匯出為 Excel
        df = pd.DataFrame(data_list)
        # 調整欄位顯示順序
        columns_order = [
            "Filename",
            "Invoice No",
            "Invoice Date",
            "PO Number",
            "From",
            "To",
            "SKU",
            "Total Amount",
        ]
        df = df.reindex(columns=columns_order)

        output_excel = os.path.join(OUTPUT_FOLDER, "invoices_parsed.xlsx")
        df.to_excel(output_excel, index=False)

        return send_file(output_excel, as_attachment=True)

    return """
    <!doctype html>
    <html>
    <head>
        <title>Invoice PDF 自動擷取工具</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 50px auto; max-width: 600px; text-align: center; }
            .drop-zone { border: 2px dashed #007bff; padding: 40px; border-radius: 8px; margin-bottom: 20px; }
            input[type="submit"] { background: #007bff; color: white; border: none; padding: 10px 20px; font-size: 16px; border-radius: 5px; cursor: pointer; }
            input[type="submit"]:hover { background: #0056b3; }
        </style>
    </head>
    <body>
        <h2>上傳發票 PDF（支援多檔批次處理）</h2>
        <form method="post" enctype="multipart/form-data">
            <div class="drop-zone">
                <input type="file" name="files" multiple accept=".pdf">
            </div>
            <input type="submit" value="開始解析並下載 Excel">
        </form>
    </body>
    </html>
    """


if __name__ == "__main__":
    app.run(debug=True, port=5000)
