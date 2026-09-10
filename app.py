import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# 網頁設定
st.set_page_config(page_title="PDF 提單與發票自動整理工具", layout="wide")
st.title("📄 PDF 提單 & 發票自動轉 Excel 工具")
st.write(
    "支援一次拖曳上傳多個 PDF 檔案，自動擷取資料並匯出為標準 Excel。"
)


def parse_pdf(file_bytes):
    """解析單一 PDF 記憶體檔案"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # 1. INVOICE NUMBER
    inv_match = re.search(
        r"Invoice\s*no\.?[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_number = inv_match.group(1).strip() if inv_match else ""

    # 2. FREIGHT AMOUNT (Total 金額)
    amount_match = re.search(
        r"Total\s*\$?\s*([0-9,]+\.[0-9]{2})", full_text, re.IGNORECASE
    )
    freight_amount = (
        f"${amount_match.group(1).strip()}" if amount_match else ""
    )

    # 3. INVOICE DATE
    date_match = re.search(
        r"Invoice\s*date[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 4. MATERIAL (SKU#)
    sku_match = re.search(
        r"SKU#?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    ) or re.search(r"(CLSAC[A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    material = sku_match.group(1).strip() if sku_match else ""

    # 5. REMARK (PO#)
    po_match = re.search(
        r"PO#?[:：]?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    po_number = po_match.group(1).strip() if po_match else ""
    remark = f"PO#: {po_number}" if po_number else ""

    # 6. FROM (例: Delta + 601 -> delta601)
    from_name = re.search(
        r"SHIP FROM[\s\S]*?Name[:：]?\s*([A-Za-z0-9]+)", full_text, re.IGNORECASE
    )
    from_addr = re.search(
        r"SHIP FROM[\s\S]*?Address[:：]?\s*(\d+)", full_text, re.IGNORECASE
    )
    from_val = (
        f"{from_name.group(1).lower()}{from_addr.group(1)}"
        if (from_name and from_addr)
        else ""
    )

    # 7. TO (例: Peak + 1991 -> peak1991)
    to_name = re.search(
        r"DELIVERY TO[\s\S]*?Name[:：]?\s*([A-Za-z0-9]+)",
        full_text,
        re.IGNORECASE,
    )
    to_addr = re.search(
        r"DELIVERY TO[\s\S]*?Address[:：]?\s*(\d+)", full_text, re.IGNORECASE
    )
    to_val = (
        f"{to_name.group(1).lower()}{to_addr.group(1)}"
        if (to_name and to_addr)
        else ""
    )

    return {
        "date": "",
        "INVOICE NUMBER": invoice_number,
        "FREIGHT AMOUNT": freight_amount,
        "INVOICE DATE": invoice_date,
        "MATERIAL": material,
        "REMARK": remark,
        "INBOUND": "",
        "NOTES": "",
        "FROM": from_val,
        "TO": to_val,
    }


# 上傳元件
uploaded_files = st.file_uploader(
    "請選擇或拖曳 PDF 檔案至此", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
    records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_pdf(file.read())
        records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    df = pd.DataFrame(records)

    st.success(f"🎉 成功解析 {len(records)} 個檔案！")

    # 預覽表格
    st.subheader("📊 擷取結果預覽")
    st.dataframe(df, use_container_width=True)

    # 轉成 Excel 下載
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="整理結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )