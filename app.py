import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="Invoice PDF 擷取工具", layout="centered")

st.title("Invoice PDF 自動擷取工具")
st.write("上傳發票 PDF（支援多檔批次處理），系統會自動解析並匯出 Excel。")


def parse_pdf(pdf_file):
    full_text = ""
    # Streamlit 的 UploadedFile 直接傳入 pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # 1. Invoice No.
    inv_no_match = re.search(
        r"Invoice\s*no\.?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_no = inv_no_match.group(1).strip() if inv_no_match else ""

    # 2. Invoice Date
    date_match = re.search(
        r"Invoice\s*date\s*[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 3. PO Number
    po_match = re.search(
        r"PO\s*#?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    po_number = po_match.group(1).strip() if po_match else ""

    # 4. Total Amount
    total_match = re.search(r"Total\s*[\$]?\s*([0-9,]+\.[0-9]{2})", full_text)
    total_amount = total_match.group(1).strip() if total_match else ""

    # 5. SKU
    sku_match = re.search(
        r"SKU\s*#?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    sku = sku_match.group(1).strip() if sku_match else ""

    # 6 & 7. FROM & TO（紅色螢光筆部分）
    from_val = ""
    to_val = ""

    route_match = re.search(
        r"from\s+(.*?)\s+(?:Plano\s+)?to\s+(.*?)(?:\s+on|\s+Dock|\n|$)",
        full_text,
        re.IGNORECASE,
    )

    if route_match:
        from_val = route_match.group(1).strip()  # 抓出: 601 Delta
        to_val = route_match.group(2).strip()  # 抓出: 1991 Peak Smart
    else:
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


# 檔案上傳元件
uploaded_files = st.file_uploader(
    "選擇或拖曳 PDF 發票至此", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
    data_list = []
    with st.spinner("正在解析 PDF 檔案..."):
        for file in uploaded_files:
            try:
                res = parse_pdf(file)
                res["Filename"] = file.name
                data_list.append(res)
            except Exception as e:
                st.error(f"解析 {file.name} 失敗: {e}")

    if data_list:
        # 轉換為 DataFrame
        df = pd.DataFrame(data_list)
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

        st.success(f"成功解析 {len(data_list)} 個檔案！")
        st.dataframe(df)

        # 輸出成 Excel 供下載
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        excel_data = output.getvalue()

        st.download_button(
            label="📥 下載 Excel 結果檔",
            data=excel_data,
            file_name="invoices_parsed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
