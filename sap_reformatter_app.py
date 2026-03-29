"""
SAP Revenue Report Reformatter — Streamlit App
Run with:  streamlit run sap_reformatter_app.py
"""

import io
import math
import re
import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SAP Revenue Reformatter",
    page_icon="📊",
    layout="centered",
)

# ── Styling ───────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    /* Tighten up the top padding */
    .block-container { padding-top: 2rem; }

    /* Upload box */
    [data-testid="stFileUploader"] {
        border: 2px dashed #d0d5dd;
        border-radius: 10px;
        padding: 1rem;
        background: #f9fafb;
    }

    /* Success / info boxes */
    .result-box {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        margin-top: 0.75rem;
    }
    .stat-row {
        display: flex;
        gap: 2rem;
        flex-wrap: wrap;
        margin-top: 0.5rem;
    }
    .stat-item {
        text-align: center;
    }
    .stat-num {
        font-size: 1.6rem;
        font-weight: 700;
        color: #15803d;
        line-height: 1.1;
    }
    .stat-label {
        font-size: 0.78rem;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
</style>
""", unsafe_allow_html=True)

# ── Constants (mirror of sap_reformatter.py) ──────────────────────────────────

OUTPUT_COLUMNS = [
    "Sales Order Number", "Item Code", "Description",
    "Sales Rep", "Sales Region", "Customer",
    "Postal Code", "Units", "Value",
]
RAW_COL_INDICES       = [2, 3, 4, 6, 8, 10, 11, 12, 13]
FORWARD_FILL_POSITIONS = [2, 3, 5, 6, 7, 8, 9, 10]
PLACEHOLDER_DESCRIPTION = "Customers"
TOTAL_ROW_MARKER        = "Overall Result"

# ── Core transform (works on a BytesIO / file-like object) ────────────────────

def find_header_row(df):
    for i in range(len(df)):
        val = df.iloc[i, 2]
        if isinstance(val, str) and "Sales Order Number" in val:
            return i
    raise ValueError(
        "Could not find the data header in this file.  "
        "Please make sure you're uploading the raw SAP BEx export (.xls)."
    )

def round_half_away(x):
    if pd.isna(x):
        return pd.NA
    f = float(x)
    return int(math.copysign(math.floor(abs(f) + 0.5), f))

def reformat_bytes(file_bytes, filename):
    """
    Accepts raw file bytes, returns (output_bytes, stats_dict).
    Raises ValueError with a user-friendly message on bad input.
    """
    buf = io.BytesIO(file_bytes)

    try:
        df_raw = pd.read_excel(buf, sheet_name="Sheet1", header=None)
    except Exception:
        raise ValueError(
            "Could not read a sheet named **Sheet1** from this file.  \n"
            "Please upload the original raw SAP export — not a previously reformatted file."
        )

    header_row = find_header_row(df_raw)

    data = df_raw.iloc[header_row + 1:].copy().reset_index(drop=True)
    raw_count = len(data)

    # Remove grand-total and placeholder rows
    mask_total = data.iloc[:, 0].astype(str).str.strip() == TOTAL_ROW_MARKER
    data = data[~mask_total].reset_index(drop=True)

    desc_pos = RAW_COL_INDICES[2]
    mask_placeholder = data.iloc[:, desc_pos] == PLACEHOLDER_DESCRIPTION
    removed_count = mask_placeholder.sum()
    data = data[~mask_placeholder].reset_index(drop=True)

    # Forward-fill merged-cell gaps
    for col_pos in FORWARD_FILL_POSITIONS:
        data.iloc[:, col_pos] = data.iloc[:, col_pos].replace("", pd.NA).ffill()

    # Postal code: pad 4-digit leading-zero codes to 5 digits
    data.iloc[:, 11] = data.iloc[:, 11].astype(str).map(
        lambda x: ("0" + x) if (re.fullmatch(r"\d{4}", x) and x.startswith("0")) else x
    )

    # Select + rename
    output = data.iloc[:, RAW_COL_INDICES].copy()
    output.columns = OUTPUT_COLUMNS

    # Round Value
    output["Value"] = (
        pd.to_numeric(output["Value"], errors="coerce")
        .map(round_half_away)
        .astype("Int64")
    )
    output["Units"] = pd.to_numeric(output["Units"], errors="coerce").astype("Int64")

    # Write to bytes buffer
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        output.to_excel(writer, sheet_name="Copy of Revenue Units & Dollars", index=False)
    out_buf.seek(0)

    # Stats
    total_value = pd.to_numeric(output["Value"], errors="coerce").sum()
    unique_orders = output["Sales Order Number"].nunique()
    stats = {
        "rows":          len(output),
        "orders":        unique_orders,
        "total_value":   total_value,
        "removed_rows":  int(removed_count) + int(mask_total.sum()),
    }

    return out_buf.getvalue(), stats

# ── UI ────────────────────────────────────────────────────────────────────────

st.title("📊 SAP Revenue Reformatter")
st.markdown(
    "Upload the raw SAP BEx export and get back a clean, analysis-ready Excel file — "
    "no manual cleanup needed."
)
st.divider()

# File uploader
uploaded = st.file_uploader(
    "Drop your raw SAP export here",
    type=["xls", "xlsx"],
    help="This should be the file that comes directly out of SAP — the one with "
         "all the messy header rows at the top.",
)

if uploaded is not None:
    file_bytes = uploaded.read()
    original_name = uploaded.name
    stem = original_name.rsplit(".", 1)[0]
    output_filename = stem + "_reformatted.xlsx"

    with st.spinner("Processing…"):
        try:
            out_bytes, stats = reformat_bytes(file_bytes, original_name)
            success = True
        except ValueError as e:
            st.error(str(e))
            success = False
        except Exception as e:
            st.error(f"Unexpected error: {e}")
            success = False

    if success:
        st.success("✅ File reformatted successfully!")

        # Stats row
        st.markdown(f"""
        <div class="result-box">
          <div class="stat-row">
            <div class="stat-item">
              <div class="stat-num">{stats['rows']:,}</div>
              <div class="stat-label">Line items</div>
            </div>
            <div class="stat-item">
              <div class="stat-num">{stats['orders']:,}</div>
              <div class="stat-label">Sales orders</div>
            </div>
            <div class="stat-item">
              <div class="stat-num">${stats['total_value']:,.0f}</div>
              <div class="stat-label">Total value</div>
            </div>
            <div class="stat-item">
              <div class="stat-num">{stats['removed_rows']:,}</div>
              <div class="stat-label">Junk rows removed</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("")  # spacer

        # Download button
        st.download_button(
            label="⬇️  Download reformatted file",
            data=out_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        # Collapsible preview
        with st.expander("👀 Preview first 20 rows"):
            preview_df = pd.read_excel(io.BytesIO(out_bytes), sheet_name=0, nrows=20)
            st.dataframe(preview_df, use_container_width=True, hide_index=True)

st.divider()

# What this does — expandable
with st.expander("ℹ️  What does this tool do?"):
    st.markdown("""
**This tool cleans up the raw SAP BEx revenue export into a tidy spreadsheet.**

The raw SAP file has ~38 rows of metadata at the top, paired code-and-description columns, 
placeholder rows, and a grand-total row — none of which you need.  
This tool removes all of that automatically and produces a clean 9-column file:

| Output column | Source |
|---|---|
| Sales Order Number | As-is from SAP |
| Item Code | "Material" field |
| Description | Material description text |
| Sales Rep | Sales rep name (drops numeric code) |
| Sales Region | Sales office description (drops office code) |
| Customer | Customer name (drops SAP customer code) |
| Postal Code | Preserved as text (leading zeros kept) |
| Units | As-is |
| Value | Dealer Net, rounded to whole dollars |

**What gets removed:**
- The ~38-row SAP metadata header
- "Customers" placeholder rows (housekeeping rows with no real transaction)
- The "Overall Result" grand-total row at the bottom

**File format:** Upload `.xls` or `.xlsx` · Output is always `.xlsx`
""")
