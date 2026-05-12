import streamlit as st
import zipfile
import re
import xml.etree.ElementTree as ET
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
from io import BytesIO
import pandas as pd

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Masterfile Comment Extractor",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Background */
.stApp {
    background: #F4F6FA;
}

/* Hide default Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }

/* Top banner */
.top-banner {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    border-radius: 16px;
    padding: 36px 40px 32px 40px;
    margin-bottom: 28px;
    box-shadow: 0 8px 32px rgba(31,78,121,0.18);
}
.top-banner h1 {
    color: #FFFFFF;
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 6px 0;
    letter-spacing: -0.5px;
}
.top-banner p {
    color: #BDD7EE;
    font-size: 1rem;
    margin: 0;
    font-weight: 400;
}

/* Cards */
.card {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 28px 32px;
    margin-bottom: 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border: 1px solid #E8EDF4;
}
.card-title {
    font-size: 0.85rem;
    font-weight: 600;
    color: #1F4E79;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 16px;
}

/* Step badges */
.step-badge {
    display: inline-block;
    background: #1F4E79;
    color: white;
    width: 26px;
    height: 26px;
    border-radius: 50%;
    text-align: center;
    line-height: 26px;
    font-size: 0.8rem;
    font-weight: 700;
    margin-right: 10px;
}

/* Legend pills */
.pill-orange {
    display: inline-block;
    background: #FDE9D9;
    border: 1.5px solid #C65911;
    color: #C65911;
    border-radius: 20px;
    padding: 3px 14px;
    font-size: 0.82rem;
    font-weight: 600;
    margin-right: 8px;
}
.pill-yellow {
    display: inline-block;
    background: #FFFFE0;
    border: 1.5px solid #7D6608;
    color: #7D6608;
    border-radius: 20px;
    padding: 3px 14px;
    font-size: 0.82rem;
    font-weight: 600;
}

/* Stat boxes */
.stat-row {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
}
.stat-box {
    flex: 1;
    background: #EEF4FB;
    border-radius: 10px;
    padding: 16px 20px;
    text-align: center;
    border: 1px solid #BDD7EE;
}
.stat-number {
    font-size: 2rem;
    font-weight: 700;
    color: #1F4E79;
    line-height: 1;
}
.stat-label {
    font-size: 0.78rem;
    color: #5A7A9A;
    font-weight: 500;
    margin-top: 4px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Preview table */
.preview-wrap {
    overflow-x: auto;
    border-radius: 8px;
    border: 1px solid #E8EDF4;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #1F4E79, #2E75B6) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 14px 32px !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    font-family: 'DM Sans', sans-serif !important;
    width: 100% !important;
    cursor: pointer !important;
    transition: opacity 0.2s !important;
    box-shadow: 0 4px 16px rgba(31,78,121,0.25) !important;
}
.stDownloadButton > button:hover {
    opacity: 0.88 !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    border: 2px dashed #BDD7EE !important;
    border-radius: 12px !important;
    background: #F8FBFF !important;
    padding: 8px !important;
}

/* Select box */
[data-testid="stSelectbox"] label {
    font-weight: 600 !important;
    color: #1F4E79 !important;
}

/* Info box */
.info-box {
    background: #EEF4FB;
    border-left: 4px solid #2E75B6;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    font-size: 0.88rem;
    color: #2E4057;
    margin-top: 12px;
}

/* Success box */
.success-box {
    background: #EBF5EB;
    border-left: 4px solid #70AD47;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-size: 0.92rem;
    color: #375623;
    font-weight: 500;
    margin: 16px 0;
}

/* Error box */
.error-box {
    background: #FDE8E8;
    border-left: 4px solid #C00000;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-size: 0.92rem;
    color: #7B0000;
    font-weight: 500;
    margin: 16px 0;
}
</style>
""", unsafe_allow_html=True)


# ── Helper: extract headers from workbook ─────────────────────────────────────
def get_headers(ws):
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[cell.column] = str(cell.value).strip()
    return headers


# ── Helper: extract comments with colour classification ───────────────────────
ORANGE_RGB = "FFFFC000"
YELLOW_RGB = "FFFFFF00"

def extract_comments(file_bytes, uid_col_name):
    """
    Returns:
        sku_comments : dict  {sku: [(attr, comment_text, ctype)]}
        error        : str | None
    """
    try:
        wb = load_workbook(BytesIO(file_bytes))
        ws = wb.active
    except Exception as e:
        return None, f"Could not open file: {e}"

    headers = get_headers(ws)
    if not headers:
        return None, "No headers found in row 1."

    # Find the UID column index
    uid_col_idx = None
    for col_idx, col_name in headers.items():
        if col_name.strip().lower() == uid_col_name.strip().lower():
            uid_col_idx = col_idx
            break
    if uid_col_idx is None:
        return None, f"Column '{uid_col_name}' not found in row 1. Please check the column name."

    # Build uid → row map
    uid_map = {}
    for row in ws.iter_rows(min_row=2):
        uid_val = row[uid_col_idx - 1].value
        if uid_val is not None:
            uid_map[row[0].row] = str(uid_val).strip()

    # Read threaded comments
    try:
        raw = BytesIO(file_bytes)
        with zipfile.ZipFile(raw) as z:
            filenames = z.namelist()
            tc_files = [f for f in filenames if 'threadedComment' in f]
            legacy_files = [f for f in filenames if f == 'xl/comments1.xml']

            if tc_files:
                with z.open(tc_files[0]) as f:
                    root = ET.parse(f).getroot()
                ns = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments'
                comment_nodes = root.findall(f'{{{ns}}}threadedComment')
                def get_ref(node): return node.get('ref')
                def get_text(node):
                    t = node.find(f'{{{ns}}}text')
                    return t.text if t is not None else ""
            elif legacy_files:
                with z.open('xl/comments1.xml') as f:
                    root = ET.parse(f).getroot()
                ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                comment_nodes = root.findall(f'.//{{{ns}}}comment')
                def get_ref(node): return node.get('ref')
                def get_text(node):
                    texts = node.findall(f'.//{{{ns}}}t')
                    return "".join(t.text or "" for t in texts)
            else:
                return None, "No comments found in this file."
    except Exception as e:
        return None, f"Error reading comments: {e}"

    sku_comments = defaultdict(list)

    for node in comment_nodes:
        ref = get_ref(node)
        text = get_text(node)
        if not ref or not text:
            continue
        m = re.match(r'([A-Z]+)(\d+)', ref)
        if not m:
            continue
        col_letter, row_str = m.groups()
        row_num = int(row_str)
        col_num = column_index_from_string(col_letter)
        if row_num == 1:
            continue

        uid = uid_map.get(row_num)
        if not uid:
            continue

        attr = headers.get(col_num, col_letter)

        cell = ws.cell(row=row_num, column=col_num)
        fg = cell.fill.fgColor
        color = fg.rgb if fg.type == 'rgb' else ""

        if color == ORANGE_RGB:
            ctype = "confirmation"
        elif color == YELLOW_RGB:
            cell_val = ws.cell(row=row_num, column=col_num).value
            if cell_val is not None and str(cell_val).strip() != "":
                ctype = "action"
            else:
                ctype = "missing"
        else:
            ctype = "other"

        sku_comments[uid].append((attr, text, ctype))

    if not sku_comments:
        return None, "No comments with recognised cell highlights found. Make sure cells are highlighted orange or yellow."

    return dict(sku_comments), None


# ── Helper: build Excel report ────────────────────────────────────────────────
def build_report(sku_comments, uid_col_name, file_name):
    out = Workbook()

    navy = PatternFill("solid", fgColor="1F4E79")
    blue = PatternFill("solid", fgColor="2E75B6")
    sku_hdr_colors = ["BDD7EE", "C6EFCE"]
    sku_row_colors = ["DEEAF1", "EBF5EB"]
    orange_hdr   = PatternFill("solid", fgColor="C65911")
    action_hdr   = PatternFill("solid", fgColor="7D6608")
    missing_hdr  = PatternFill("solid", fgColor="C00000")
    other_hdr    = PatternFill("solid", fgColor="595959")
    orange_cell  = PatternFill("solid", fgColor="FDE9D9")
    action_cell  = PatternFill("solid", fgColor="FFFFE0")
    missing_cell = PatternFill("solid", fgColor="FFE0E0")
    other_cell   = PatternFill("solid", fgColor="EDEDED")

    thin  = Side(style="thin", color="B8CCE4")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr   = Alignment(horizontal="center", vertical="center")
    left  = Alignment(horizontal="left", vertical="center", indent=1)
    nowrap = Alignment(vertical="center", wrap_text=False)

    def hdr_cell(ws, row, col, val, fill=None, color="FFFFFF"):
        fill = fill or navy
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", bold=True, color=color, size=11)
        c.fill = fill
        c.alignment = ctr
        c.border = bdr

    def data_cell(ws, row, col, val, fill, bold=False, align=None, size=10):
        align = align or left
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name="Arial", bold=bold, size=size)
        c.fill = fill
        c.alignment = align
        c.border = bdr

    total_comments = sum(len(v) for v in sku_comments.values())
    conf_count    = sum(1 for v in sku_comments.values() for _,_,t in v if t == "confirmation")
    action_count  = sum(1 for v in sku_comments.values() for _,_,t in v if t == "action")
    missing_count = sum(1 for v in sku_comments.values() for _,_,t in v if t == "missing")
    other_count   = sum(1 for v in sku_comments.values() for _,_,t in v if t == "other")
    source_name = file_name.replace(".xlsx", "").replace("_", " ")

    # ── Sheet 1: Summary ──────────────────────────────────────────────────────
    ws_sum = out.active
    ws_sum.title = "Summary"
    ws_sum.column_dimensions["A"].width = 6
    ws_sum.column_dimensions["B"].width = 34
    ws_sum.column_dimensions["C"].width = 20
    ws_sum.column_dimensions["D"].width = 42

    ws_sum.merge_cells("A1:D1")
    c = ws_sum["A1"]
    c.value = f"MASTERFILE COMMENT REPORT  —  {source_name}"
    c.font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    c.fill = navy
    c.alignment = Alignment(horizontal="center", vertical="center")

    ws_sum.merge_cells("A2:D2")
    c = ws_sum["A2"]
    c.value = (
        f"Total {uid_col_name}s: {len(sku_comments)}"
        f"    |    Total Comments: {total_comments}"
        f"    |    Confirmation: {conf_count}"
        f"    |    Action Required: {action_count}"
        f"    |    Data Missing: {missing_count}"
        f"    |    Other: {other_count}"
    )
    c.font = Font(name="Arial", size=10, color="FFFFFF")
    c.fill = blue
    c.alignment = Alignment(horizontal="center", vertical="center")

    for col, val in enumerate(["#", uid_col_name, "Total Comments", "Attributes Affected"], start=1):
        hdr_cell(ws_sum, 3, col, val)

    for i, (sku, comments) in enumerate(sku_comments.items(), start=1):
        r = 3 + i
        fill = PatternFill("solid", fgColor=sku_row_colors[i % 2])
        attrs = ", ".join(dict.fromkeys(a for a, _, __ in comments))
        data_cell(ws_sum, r, 1, i,             fill=fill, align=ctr)
        data_cell(ws_sum, r, 2, sku,           fill=fill, bold=True)
        data_cell(ws_sum, r, 3, len(comments), fill=fill, align=ctr)
        data_cell(ws_sum, r, 4, attrs,         fill=fill)

    # ── Sheet 2: Comment Report ───────────────────────────────────────────────
    ws_rpt = out.create_sheet("Comment Report")
    ws_rpt.column_dimensions["A"].width = 30
    ws_rpt.column_dimensions["B"].width = 26
    ws_rpt.column_dimensions["C"].width = 40
    ws_rpt.column_dimensions["D"].width = 40
    ws_rpt.column_dimensions["E"].width = 40
    ws_rpt.column_dimensions["F"].width = 40

    ws_rpt.merge_cells("A1:F1")
    c = ws_rpt["A1"]
    c.value = "All Comments — Attribute Level Detail"
    c.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill = navy
    c.alignment = Alignment(horizontal="center", vertical="center")

    hdr_cell(ws_rpt, 2, 1, uid_col_name)
    hdr_cell(ws_rpt, 2, 2, "Attribute")
    hdr_cell(ws_rpt, 2, 3, "Confirmation Comment", fill=orange_hdr)
    hdr_cell(ws_rpt, 2, 4, "Action Required",       fill=action_hdr)
    hdr_cell(ws_rpt, 2, 5, "Data Missing",           fill=missing_hdr)
    hdr_cell(ws_rpt, 2, 6, "Other Comments",         fill=other_hdr)

    row_idx = 3
    for i, (sku, comments) in enumerate(sku_comments.items()):
        hdr_fill = PatternFill("solid", fgColor=sku_hdr_colors[i % 2])
        row_fill = PatternFill("solid", fgColor=sku_row_colors[i % 2])

        for col in range(1, 7):
            c = ws_rpt.cell(row=row_idx, column=col)
            c.fill = hdr_fill
            c.font = Font(name="Arial", bold=True, color="000000", size=10)
            c.border = bdr
            c.alignment = left
        ws_rpt.cell(row=row_idx, column=1).value = sku
        ws_rpt.cell(row=row_idx, column=2).value = f"▼  {len(comments)} comment(s)"
        row_idx += 1

        attr_map = defaultdict(lambda: {"confirmation": "", "action": "", "missing": "", "other": ""})
        for attr, comment, ctype in comments:
            attr_map[attr][ctype] = comment

        for attr, vals in attr_map.items():
            data_cell(ws_rpt, row_idx, 1, sku,  fill=row_fill)
            data_cell(ws_rpt, row_idx, 2, attr, fill=row_fill, bold=True)

            cf = orange_cell if vals["confirmation"] else row_fill
            c = ws_rpt.cell(row=row_idx, column=3, value=vals["confirmation"] or "-")
            c.font = Font(name="Arial", size=10); c.fill = cf; c.alignment = nowrap; c.border = bdr

            af = action_cell if vals["action"] else row_fill
            c = ws_rpt.cell(row=row_idx, column=4, value=vals["action"] or "-")
            c.font = Font(name="Arial", size=10); c.fill = af; c.alignment = nowrap; c.border = bdr

            mf = missing_cell if vals["missing"] else row_fill
            c = ws_rpt.cell(row=row_idx, column=5, value=vals["missing"] or "-")
            c.font = Font(name="Arial", size=10); c.fill = mf; c.alignment = nowrap; c.border = bdr

            of = other_cell if vals["other"] else row_fill
            c = ws_rpt.cell(row=row_idx, column=6, value=vals["other"] or "-")
            c.font = Font(name="Arial", size=10); c.fill = of; c.alignment = nowrap; c.border = bdr

            row_idx += 1

    ws_rpt.freeze_panes = "A3"
    ws_rpt.auto_filter.ref = f"A2:F{row_idx - 1}"

    buf = BytesIO()
    out.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="top-banner">
    <h1>📋 Masterfile Comment Extractor</h1>
    <p>Upload any Masterfile → choose your Unique ID column → download a clean comment report instantly.</p>
</div>
""", unsafe_allow_html=True)

# Legend
st.markdown("""
<div class="card">
    <div class="card-title">Comment Type Legend</div>
    <span class="pill-orange">🟠 Orange highlight → Confirmation Comment</span>
    <span class="pill-yellow">🟡 Yellow + has data → Action Required</span>
    <span style="display:inline-block;background:#FFE0E0;border:1.5px solid #C00000;color:#C00000;border-radius:20px;padding:3px 14px;font-size:0.82rem;font-weight:600;margin-right:8px;">🔴 Yellow + empty cell → Data Missing</span>
    <span style="display:inline-block;background:#EDEDED;border:1.5px solid #595959;color:#595959;border-radius:20px;padding:3px 14px;font-size:0.82rem;font-weight:600;">⚪ Other colour → Other Comments</span>
    <div class="info-box" style="margin-top:14px;">
        <b>Orange</b> → Confirmation needed (sign-off required from brand/EPM)<br>
        <b>Yellow + data present</b> → Action Required (e.g. value incorrect, exceeds limit)<br>
        <b>Yellow + cell empty</b> → Data Missing (mandatory field not filled)<br>
        <b>Any other colour</b> → Other Comments (captured separately for review)
    </div>
</div>
""", unsafe_allow_html=True)

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.markdown('<div class="card"><div class="card-title"><span class="step-badge">1</span>Upload Masterfile</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file:
    file_bytes = uploaded_file.read()

    # Peek at headers for dropdown
    try:
        wb_peek = load_workbook(BytesIO(file_bytes))
        ws_peek = wb_peek.active
        col_names = [
            str(cell.value).strip()
            for cell in ws_peek[1]
            if cell.value is not None
        ]
    except Exception:
        col_names = []

    # ── Step 2: Choose UID column ─────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title"><span class="step-badge">2</span>Select Unique ID Column</div>', unsafe_allow_html=True)

    if col_names:
        uid_col = st.selectbox(
            "Choose the column that uniquely identifies each product row:",
            options=col_names,
            index=0,
        )
        st.markdown(
            f'<div class="info-box">✅ Using <b>{uid_col}</b> as the Unique ID. '
            f'Each comment in the report will be grouped under its {uid_col} value.</div>',
            unsafe_allow_html=True
        )
    else:
        uid_col = st.text_input("Enter column name exactly as it appears in row 1 of your file:", value="SKU")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 3: Generate ──────────────────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title"><span class="step-badge">3</span>Generate Report</div>', unsafe_allow_html=True)

    if st.button("⚡ Extract Comments & Build Report", use_container_width=True):
        with st.spinner("Reading comments and building your report..."):
            sku_comments, error = extract_comments(file_bytes, uid_col)

        if error:
            st.markdown(f'<div class="error-box">❌ {error}</div>', unsafe_allow_html=True)
        else:
            total_skus     = len(sku_comments)
            total_comments = sum(len(v) for v in sku_comments.values())
            conf_count     = sum(1 for v in sku_comments.values() for _, _, t in v if t == "confirmation")
            action_count   = sum(1 for v in sku_comments.values() for _, _, t in v if t == "action")
            missing_count  = sum(1 for v in sku_comments.values() for _, _, t in v if t == "missing")
            other_count    = sum(1 for v in sku_comments.values() for _, _, t in v if t == "other")

            st.markdown(f'<div class="success-box">✅ Found <b>{total_comments}</b> comments across <b>{total_skus}</b> {uid_col}(s). Report is ready!</div>', unsafe_allow_html=True)

            # Stat boxes
            st.markdown(f"""
            <div class="stat-row">
                <div class="stat-box">
                    <div class="stat-number">{total_skus}</div>
                    <div class="stat-label">{uid_col}s</div>
                </div>
                <div class="stat-box">
                    <div class="stat-number">{total_comments}</div>
                    <div class="stat-label">Total Comments</div>
                </div>
                <div class="stat-box" style="background:#FDE9D9; border-color:#C65911;">
                    <div class="stat-number" style="color:#C65911;">{conf_count}</div>
                    <div class="stat-label" style="color:#C65911;">Confirmation</div>
                </div>
                <div class="stat-box" style="background:#FFFFE0; border-color:#7D6608;">
                    <div class="stat-number" style="color:#7D6608;">{action_count}</div>
                    <div class="stat-label" style="color:#7D6608;">Action Required</div>
                </div>
                <div class="stat-box" style="background:#FFE0E0; border-color:#C00000;">
                    <div class="stat-number" style="color:#C00000;">{missing_count}</div>
                    <div class="stat-label" style="color:#C00000;">Data Missing</div>
                </div>
                <div class="stat-box" style="background:#EDEDED; border-color:#595959;">
                    <div class="stat-number" style="color:#595959;">{other_count}</div>
                    <div class="stat-label" style="color:#595959;">Other</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Preview table
            st.markdown("**Preview — first 10 rows of Comment Report:**")
            preview_rows = []
            for sku, comments in list(sku_comments.items())[:5]:
                attr_map = defaultdict(lambda: {"confirmation": "", "action": "", "missing": "", "other": ""})
                for attr, comment, ctype in comments:
                    attr_map[attr][ctype] = comment
                for attr, vals in list(attr_map.items())[:3]:
                    preview_rows.append({
                        uid_col: sku,
                        "Attribute": attr,
                        "Confirmation Comment": vals["confirmation"] or "-",
                        "Action Required":      vals["action"] or "-",
                        "Data Missing":         vals["missing"] or "-",
                        "Other Comments":       vals["other"] or "-",
                    })

            df_preview = pd.DataFrame(preview_rows[:10])
            st.dataframe(df_preview, use_container_width=True, hide_index=True)

            # Build & offer download
            report_buf = build_report(sku_comments, uid_col, uploaded_file.name)
            out_name = uploaded_file.name.replace(".xlsx", "") + "_Comment_Report.xlsx"

            st.download_button(
                label="⬇️  Download Excel Report",
                data=report_buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.markdown("""
    <div style="text-align:center; padding: 40px 0; color:#8BA4C0;">
        <div style="font-size:3rem;">📂</div>
        <div style="font-size:1rem; font-weight:500; margin-top:8px;">Upload your Masterfile (.xlsx) above to get started</div>
        <div style="font-size:0.85rem; margin-top:4px;">Supports any marketplace Masterfile — Amazon, eBay, OTTO, CDiscount, MediaMarkt</div>
    </div>
    """, unsafe_allow_html=True)
