
import os
import re
from io import BytesIO

import pandas as pd
import streamlit as st


# ==========================================
# PAGE CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="Sales Summary Dashboard",
    page_icon="📊",
    layout="wide"
)


# ==========================================
# HELPER FUNCTIONS
# ==========================================

def clean_order_ids(series):
    """Normalize order IDs to improve return matching."""
    values = series.astype("string").str.strip()

    # Handle Excel IDs that were converted to floats, e.g. 12345.0
    values = values.str.replace(r"\.0$", "", regex=True)

    return values.mask(values.eq(""), pd.NA)


def clean_states(series):
    """Normalize state names so both platforms can be combined."""
    values = series.astype("string").str.strip().str.title()

    return values.mask(values.eq(""), "Unknown").fillna("Unknown")


def read_jiomart_csv(uploaded_file):
    """Read a JioMart CSV with common encoding fallbacks."""
    uploaded_file.seek(0)

    try:
        return pd.read_csv(uploaded_file, encoding="utf-8-sig")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, encoding="cp1252")


def require_columns(df, columns, file_name):
    """Validate the uploaded file's column names."""
    missing = [col for col in columns if col not in df.columns]

    if missing:
        raise ValueError(
            f"{file_name} is missing these columns: "
            + ", ".join(missing)
        )


def make_output_filename(name):
    """Create a safe Excel filename."""
    name = name.strip()

    # Prevent path traversal and invalid Windows filename characters.
    name = os.path.basename(name)
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.rstrip(" .")

    if not name:
        name = "Combined_State_Wise_Summary"

    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"

    return name


def generate_summary(jiomart_file, sales_file, returns_file):
    """Process both platforms and generate a combined summary."""

    # --------------------------------------
    # 1. READ INPUT FILES
    # --------------------------------------

    jiomart_df = read_jiomart_csv(jiomart_file)
    sales_df = pd.read_excel(sales_file)
    returns_df = pd.read_excel(returns_file)

    # --------------------------------------
    # 2. DEFINE COLUMN NAMES
    # --------------------------------------

    jiomart_state = "Customer's Delivery State"
    jiomart_status = "Order Status"
    jiomart_taxable = (
        "Taxable Value (Final Invoice Amount -Taxes)"
    )

    meesho_order = "sub_order_num"
    meesho_state = "end_customer_state_new"
    meesho_taxable = "total_taxable_sale_value"

    # --------------------------------------
    # 3. VALIDATE COLUMNS
    # --------------------------------------

    require_columns(
        jiomart_df,
        [jiomart_state, jiomart_status, jiomart_taxable],
        "JioMart CSV"
    )

    require_columns(
        sales_df,
        [meesho_order, meesho_state, meesho_taxable],
        "Meesho sales file"
    )

    require_columns(
        returns_df,
        [meesho_order],
        "Meesho return file"
    )

    # --------------------------------------
    # 4. PROCESS JIOMART
    # --------------------------------------

    jiomart_df[jiomart_state] = clean_states(
        jiomart_df[jiomart_state]
    )

    jiomart_df[jiomart_status] = (
        jiomart_df[jiomart_status]
        .astype("string")
        .str.strip()
        .str.casefold()
    )

    jiomart_df[jiomart_taxable] = pd.to_numeric(
        jiomart_df[jiomart_taxable],
        errors="coerce"
    ).fillna(0)

    # Only delivered orders
    jiomart_df = jiomart_df[
        jiomart_df[jiomart_status] == "delivered"
    ].copy()

    jiomart_df = jiomart_df.rename(
        columns={
            jiomart_state: "State",
            jiomart_taxable: "Taxable_Amount"
        }
    )

    jiomart_summary = (
        jiomart_df
        .groupby("State")
        .agg(
            JioMart_Order_Count=("State", "size"),
            JioMart_Taxable_Amount=("Taxable_Amount", "sum")
        )
    )

    # --------------------------------------
    # 5. PROCESS MEESHO
    # --------------------------------------

    sales_df[meesho_order] = clean_order_ids(
        sales_df[meesho_order]
    )

    returns_df[meesho_order] = clean_order_ids(
        returns_df[meesho_order]
    )

    sales_df[meesho_state] = clean_states(
        sales_df[meesho_state]
    )

    sales_df[meesho_taxable] = pd.to_numeric(
        sales_df[meesho_taxable],
        errors="coerce"
    ).fillna(0)

    # Identify returned sub-orders
    returned_orders = set(
        returns_df[meesho_order].dropna()
    )

    # Remove sales whose sub-order ID appears in returns
    filtered_sales = sales_df[
        ~sales_df[meesho_order].isin(returned_orders)
    ].copy()

    filtered_sales = filtered_sales.rename(
        columns={
            meesho_state: "State",
            meesho_taxable: "Taxable_Amount"
        }
    )

    meesho_summary = (
        filtered_sales
        .groupby("State")
        .agg(
            Meesho_Order_Count=("State", "size"),
            Meesho_Taxable_Amount=("Taxable_Amount", "sum")
        )
    )

    # --------------------------------------
    # 6. COMBINE BOTH PLATFORMS
    # --------------------------------------

    combined = jiomart_summary.join(
        meesho_summary,
        how="outer"
    ).fillna(0)

    count_columns = [
        "JioMart_Order_Count",
        "Meesho_Order_Count"
    ]

    combined[count_columns] = (
        combined[count_columns].astype(int)
    )

    combined["Order_Count"] = (
        combined["JioMart_Order_Count"]
        + combined["Meesho_Order_Count"]
    )

    combined["Total_Taxable_Amount"] = (
        combined["JioMart_Taxable_Amount"]
        + combined["Meesho_Taxable_Amount"]
    )

    combined = combined.sort_values(
        "Total_Taxable_Amount",
        ascending=False
    )

    # Return State as a normal Excel column
    combined = combined.reset_index()

    output_columns = [
        "State",
        "JioMart_Order_Count",
        "JioMart_Taxable_Amount",
        "Meesho_Order_Count",
        "Meesho_Taxable_Amount",
        "Order_Count",
        "Total_Taxable_Amount"
    ]

    combined = combined[output_columns]

    # --------------------------------------
    # 7. CREATE EXCEL FILE IN MEMORY
    # --------------------------------------

    output_buffer = BytesIO()

    with pd.ExcelWriter(
        output_buffer,
        engine="openpyxl"
    ) as writer:
        combined.to_excel(
            writer,
            sheet_name="State Wise Summary",
            index=False
        )

    output_buffer.seek(0)

    # Additional information for the interface
    info = {
        "jiomart_rows": len(jiomart_df),
        "meesho_rows": len(filtered_sales),
        "returned_ids": len(returned_orders),
        "states": len(combined)
    }

    return combined, output_buffer.getvalue(), info


# ==========================================
# INTERFACE DESIGN
# ==========================================

st.title("📊 Sales Summary Dashboard")

st.write(
    "Combine JioMart and Meesho sales into one "
    "state-wise Excel report."
)

st.divider()

st.subheader("1. Upload your files")

left, right = st.columns(2)

with left:
    jiomart_file = st.file_uploader(
        "JioMart Invoice (CSV)",
        type=["csv"],
        help="Upload your monthly JioMart invoice CSV."
    )

    sales_file = st.file_uploader(
        "Meesho Sales File (Excel)",
        type=["xlsx", "xlsm"],
        help="Upload tcs_sales.xlsx."
    )

with right:
    returns_file = st.file_uploader(
        "Meesho Sales Return File (Excel)",
        type=["xlsx", "xlsm"],
        help="Upload tcs_sales_return.xlsx."
    )

    st.info(
        "JioMart includes delivered orders only. "
        "Meesho sales matching returned sub-orders "
        "are excluded."
    )

st.divider()

st.subheader("2. Output settings")

output_name = st.text_input(
    "Enter output filename",
    value="Combined_JioMart_Meesho_State_Wise_Summary.xlsx",
    help="The .xlsx extension is added automatically if omitted."
)

st.caption(
    "Your Excel report will contain platform-wise counts, "
    "platform-wise taxable amounts, and combined totals."
)

st.divider()

# ==========================================
# GENERATE REPORT
# ==========================================

if st.button(
    "Generate State-Wise Summary",
    type="primary",
    use_container_width=True
):
    if not jiomart_file or not sales_file or not returns_file:
        st.error("Please upload all three files first.")

    elif not output_name.strip():
        st.error("Please enter an output filename.")

    else:
        try:
            with st.spinner("Processing your sales files..."):
                summary, excel_data, info = generate_summary(
                    jiomart_file,
                    sales_file,
                    returns_file
                )

                final_filename = make_output_filename(output_name)

                # Save report details for display and download
                st.session_state["summary"] = summary
                st.session_state["excel_data"] = excel_data
                st.session_state["filename"] = final_filename
                st.session_state["info"] = info

            st.success("Your state-wise report is ready!")

        except Exception as error:
            st.error(f"Unable to generate the report: {error}")


# ==========================================
# RESULTS AND DOWNLOAD
# ==========================================

if "summary" in st.session_state:
    summary = st.session_state["summary"]
    info = st.session_state["info"]

    st.divider()
    st.subheader("3. Report preview")

    metric1, metric2, metric3 = st.columns(3)

    metric1.metric(
        "Total Orders",
        f"{int(summary['Order_Count'].sum()):,}"
    )

    metric2.metric(
        "Total Taxable Amount",
        f"₹{summary['Total_Taxable_Amount'].sum():,.2f}"
    )

    metric3.metric(
        "States Covered",
        f"{info['states']:,}"
    )

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        f"JioMart delivered rows: {info['jiomart_rows']:,} | "
        f"Meesho remaining rows: {info['meesho_rows']:,} | "
        f"Distinct return IDs: {info['returned_ids']:,}"
    )

    st.download_button(
        label="⬇️ Download Excel Report",
        data=st.session_state["excel_data"],
        file_name=st.session_state["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )

