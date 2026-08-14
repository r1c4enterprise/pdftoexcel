"""
PDF to Excel Converter — Streamlit App
Upload laporan keuangan PDF, download hasilnya sebagai Excel.
"""

import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import io


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📄",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 680px; margin: auto; }
    .title { text-align: center; font-size: 2rem; font-weight: 700; color: #1a2340; }
    .subtitle { text-align: center; color: #6b7280; margin-bottom: 2rem; }
    .stDownloadButton > button {
        width: 100%;
        background-color: #4472C4;
        color: white;
        font-weight: 600;
        font-size: 1rem;
        border-radius: 10px;
        padding: 0.6rem;
        border: none;
    }
    .stDownloadButton > button:hover { background-color: #2F5496; }
</style>
""", unsafe_allow_html=True)


# ── Helper functions ──────────────────────────────────────────────────────────

def try_parse_number(value):
    if not isinstance(value, str):
        return value
    v = value.strip()
    if v in ("", "-"):
        return value
    negative = False
    if v.startswith("(") and v.endswith(")"):
        v = v[1:-1]
        negative = True
    elif v.startswith("-"):
        v = v[1:]
        negative = True
    v = re.sub(r"[Rp$€£¥\s]", "", v)
    if re.match(r"^\d{1,3}(\.\d{3})*(,\d+)?$", v):
        v = v.replace(".", "").replace(",", ".")
    elif re.match(r"^\d{1,3}(,\d{3})*(\.\d+)?$", v):
        v = v.replace(",", "")
    elif re.match(r"^\d+(,\d+)$", v):
        v = v.replace(",", ".")
    elif not re.match(r"^\d+(\.\d+)?$", v):
        return value
    try:
        num = float(v)
        if negative:
            num = -num
        return int(num) if num == int(num) else num
    except ValueError:
        return value


def extract_tables(pdf_bytes):
    results = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                continue
            for tbl_idx, table in enumerate(tables, start=1):
                if not table or len(table) < 2:
                    continue
                header = [str(h).strip() if h else "" for h in table[0]]
                rows   = [[str(c).strip() if c else "" for c in row] for row in table[1:]]
                seen   = {}
                clean_header = []
                for col in header:
                    if col in seen:
                        seen[col] += 1
                        clean_header.append(f"{col}_{seen[col]}")
                    else:
                        seen[col] = 0
                        clean_header.append(col)
                rows = [r for r in rows if any(c.strip() for c in r)]
                parsed_rows = [[try_parse_number(c) for c in row] for row in rows]
                results.append({
                    "page": page_num,
                    "table_index": tbl_idx,
                    "headers": clean_header,
                    "rows": parsed_rows,
                })
    return results


def build_excel(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        current_row = 0
        for item in tables:
            df = pd.DataFrame(item["rows"], columns=item["headers"])
            label_df = pd.DataFrame([[f"[ Halaman {item['page']} — Tabel {item['table_index']} ]"]])
            label_df.to_excel(writer, sheet_name="Laporan Keuangan",
                              startrow=current_row, index=False, header=False)
            current_row += 1
            df.to_excel(writer, sheet_name="Laporan Keuangan",
                        startrow=current_row, index=False, header=True)
            current_row += len(df) + 1 + 2
    output.seek(0)
    wb = load_workbook(output)
    ws = wb["Laporan Keuangan"]
    _apply_styles(ws)
    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.read()


def _apply_styles(ws):
    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    label_rows  = set()
    header_rows = set()
    for row in ws.iter_rows():
        cell = row[0]
        if cell.value and str(cell.value).startswith("[ Halaman"):
            label_rows.add(cell.row)
    for r in label_rows:
        header_rows.add(r + 1)
    data_row_counters = {}
    for row in ws.iter_rows():
        row_num   = row[0].row
        is_label  = row_num in label_rows
        is_header = row_num in header_rows
        if not is_label and not is_header:
            block_start = max((h for h in header_rows if h < row_num), default=0)
            data_row_counters.setdefault(block_start, 0)
            data_row_counters[block_start] += 1
            is_even = (data_row_counters[block_start] % 2 == 0)
        else:
            is_even = False
        for cell in row:
            if is_label:
                cell.font      = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
                cell.fill      = PatternFill("solid", fgColor="2F5496")
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif is_header:
                cell.font      = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
                cell.fill      = PatternFill("solid", fgColor="4472C4")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border    = border
            else:
                cell.font = Font(size=10, name="Calibri")
                cell.fill = PatternFill("solid", fgColor="D9E1F2" if is_even else "FFFFFF")
                cell.border = border
                if isinstance(cell.value, (int, float)):
                    cell.alignment     = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '#,##0.00' if isinstance(cell.value, float) and cell.value != int(cell.value) else '#,##0'
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len    = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)
    for row in ws.row_dimensions.values():
        row.height = 18


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown('<div class="title">📄 PDF to Excel</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload laporan keuangan PDF, download hasilnya sebagai Excel.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Pilih file PDF",
    type=["pdf"],
    help="Maksimal 200MB"
)

if uploaded_file:
    st.info(f"📎 **{uploaded_file.name}** ({uploaded_file.size / 1024 / 1024:.2f} MB)")

    if st.button("⚡ Konversi ke Excel", use_container_width=True):
        with st.spinner("Memproses PDF... mohon tunggu"):
            try:
                pdf_bytes = uploaded_file.read()
                tables    = extract_tables(pdf_bytes)

                if not tables:
                    st.error("❌ Tidak ada tabel yang ditemukan di PDF ini.")
                else:
                    excel_bytes = build_excel(tables)
                    output_name = uploaded_file.name.replace(".pdf", ".xlsx")

                    st.success(f"✅ Berhasil! Ditemukan **{len(tables)} tabel** dari PDF.")

                    st.download_button(
                        label="📥 Download Excel",
                        data=excel_bytes,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {str(e)}")
