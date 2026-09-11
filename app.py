import io
import os
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="Invoice PDF 擷取工具", layout="wide")

st.title("Invoice PDF 自動擷取工具")
st.write("上傳發票 PDF（支援多檔批次處理），系統會自動解析並匯出 Excel。")


def parse_pdf(pdf_file):
    full_text = ""
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

    # 3. PO Number (修復：使用單詞邊界 \b，排除一般英文單字，至少 5 碼以上編號)
    po_match = re.search(
        r"\bPO\s*#?\s*[:：]?\s*([A-Z0-9\-]{5,})", full_text, re.IGNORECASE
    )
    po_number = ""
    if po_match:
        po_number = po_match.group(1).strip()
    else:
        # 備援：若上方沒明確標註 PO#，比對內文中的常見 PO 前綴編號 (例如 PSDELT...)
        fallback_po = re.search(r"\b(PSDELT[0-9A-Za-z]+)\b", full_text)
        if fallback_po:
            po_number = fallback_po.group(1).strip()

    # 4. Total Amount
    total_match = re.search(r"Total\s*[\$]?\s*([0-9,]+\.[0-9]{2})", full_text)
    total_amount = total_match.group(1).strip() if total_match else ""

    # 5. SKU
    sku_match = re.search(
        r"SKU\s*#?\s*[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    sku = sku_match.group(1).strip() if sku_match else ""

    # 6 & 7. FROM & TO（精準擷取）
    from_val = ""
    to_val = ""

    # 情況 A：明細描述中含有 "... from 601 Delta Plano to 1991 Peak Smart on 9/3/2026 ..."
    route_match = re.search(
        r"\bfrom\s+(.*?)\s+(?:Plano\s+)?to\s+(.*?)(?:\s+on|\s+Dock|\n|$)",
        full_text,
        re.IGNORECASE,
    )

    if route_match:
        from_val = route_match.group(1).strip()  # 601 Delta
        to_val = route_match.group(2).strip()  # 1991 Peak Smart
    else:
        # 情況 B：備援規則，若發票明細沒有，嘗試從地址或標準欄位抓取
        f_match = re.search(
            r"(?:Shipper|Ship\s*From|From)[:：]?\s*([0-9]+\s+[A-Za-z]+)",
            full_text,
            re.IGNORECASE,
        )
        t_match = re.search(
            r"(?:Consignee|Delivery\s*To|To)[:：]?\s*([0-9]+\s+[A-Za-z\s]+)",
            full_text,
            re.IGNORECASE,
        )
        if f_match:
            from_val = f_match.group(1).strip()
        if t_match:
            to_val = t_match.group(1).strip()

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
        # 顯示表格
        st.dataframe(df, use_container_width=True)

        # 匯出 Excel
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
