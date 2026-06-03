from __future__ import annotations

import gzip
import json
import subprocess
import sys
import tempfile
from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st


APP_ROOT = Path(__file__).resolve().parent
GENERATOR = APP_ROOT / "Generate-SalesBrandReport.py"
if not GENERATOR.exists():
    GENERATOR = APP_ROOT / "Code" / "Generate-SalesBrandReport.py"
REPORT_DIR = APP_ROOT / "reports"
REPORT_DIR.mkdir(exist_ok=True)


def money(value: float) -> str:
    return f"${value:,.2f}"


def number(value: float) -> str:
    return f"{value:,.0f}"


def percent(value: float) -> str:
    return f"{value:.1%}"


def safe_filename(value: str) -> str:
    invalid_chars = '<>:"/\\|?*'
    cleaned = "".join("_" if char in invalid_chars else char for char in value.strip())
    cleaned = " ".join(cleaned.split())
    return cleaned or "Nakama Sales Report"


@st.cache_data(show_spinner=False)
def load_dashboard_data(path: str) -> dict:
    with gzip.open(path, "rt", encoding="utf-8") as file:
        return json.load(file)


def save_uploaded_file(uploaded_file, folder: Path) -> Path:
    path = folder / Path(uploaded_file["name"]).name
    path.write_bytes(uploaded_file["bytes"])
    return path


def combine_sales_files(sales_files, output_path: Path) -> Path:
    frames = []
    for sales_file in sales_files:
        frames.append(pd.read_excel(sales_file, dtype={"SKU": "string"}))
    pd.concat(frames, ignore_index=True).to_excel(output_path, index=False)
    return output_path


def uploaded_file_state(uploaded_file) -> dict:
    return {
        "name": uploaded_file.name,
        "bytes": uploaded_file.getvalue(),
    }


def saved_dashboard_label(path: Path) -> str:
    if path.name.endswith(".json.gz"):
        return path.name.removesuffix(".json.gz")
    return path.stem


def saved_dashboard_files() -> list[Path]:
    return sorted(
        REPORT_DIR.glob("*.json.gz"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def companion_dashboard_files(path: Path) -> list[Path]:
    if path.name.endswith(".json.gz"):
        base_name = path.name.removesuffix(".json.gz")
    else:
        base_name = path.stem

    files = [REPORT_DIR / f"{base_name}.html", REPORT_DIR / f"{base_name}.json.gz"]
    return [file for file in files if file.exists()]


def delete_saved_dashboard(path: Path) -> None:
    for file in companion_dashboard_files(path):
        file.unlink()
    load_dashboard_data.clear()
    st.success(f"Deleted dashboard: {saved_dashboard_label(path)}")
    st.rerun()


def render_card_grid(cards: list[tuple[str, str]], class_name: str) -> None:
    html = "".join(
        f'<div class="{class_name}">'
        f'<div class="label">{escape(label)}</div>'
        f'<div class="value">{escape(value)}</div>'
        "</div>"
        for label, value in cards
    )
    st.markdown(f'<div class="{class_name}s">{html}</div>', unsafe_allow_html=True)


def render_bar_panel(
    title: str,
    frame: pd.DataFrame,
    label_column: str,
    value_column: str,
    formatter,
    color_class: str = "",
) -> None:
    if frame.empty or value_column not in frame:
        st.markdown(
            f'<div class="chart-card"><h3>{escape(title)}</h3><p>No data available.</p></div>',
            unsafe_allow_html=True,
        )
        return

    rows = frame.head(10).copy()
    max_value = float(rows[value_column].max()) if not rows.empty else 0
    if max_value <= 0:
        max_value = 1

    bars = []
    for _, row in rows.iterrows():
        label = escape(str(row.get(label_column, "")))
        value = float(row.get(value_column, 0) or 0)
        width = max(3, min(100, value / max_value * 100))
        bars.append(
            '<div class="bar-row">'
            f'<div class="bar-label">{label}</div>'
            f'<div class="bar-track"><div class="bar-fill {color_class}" style="width:{width:.2f}%"></div></div>'
            f'<div class="bar-value">{escape(formatter(value))}</div>'
            "</div>"
        )
    st.markdown(
        f'<div class="chart-card"><h3>{escape(title)}</h3>{"".join(bars)}</div>',
        unsafe_allow_html=True,
    )


def render_dashboard(report_data: dict, key_prefix: str) -> None:
    metadata = report_data.get("metadata", {})
    kpis = report_data.get("kpis", {})

    brand_df = pd.DataFrame(report_data.get("brands", []))
    customer_df = pd.DataFrame(report_data.get("customers", []))
    product_df = pd.DataFrame(report_data.get("products", []))
    customer_detail_df = pd.DataFrame(report_data.get("customerDetails", []))
    product_monthly_df = pd.DataFrame(report_data.get("productMonthlyDetails", []))
    product_matrix_df = pd.DataFrame(report_data.get("productMonthlyMatrix", []))

    st.markdown(
        '<div class="report-header">'
        f'<h2>{escape(str(metadata.get("reportTitle", "Nakama Sales Report")))}</h2>'
        f'<p>{escape(str(kpis.get("dateRange", "")))}</p>'
        "</div>",
        unsafe_allow_html=True,
    )
    render_card_grid(
        [
            ("Top Brand", str(kpis.get("topBrand", ""))),
            ("Top Customer", str(kpis.get("topCustomer", ""))),
            ("Top 10 Brand Share", percent(float(kpis.get("top10BrandShare", 0) or 0))),
        ],
        "highlight",
    )
    render_card_grid(
        [
            ("Total Sales", money(float(kpis.get("totalSales", 0) or 0))),
            ("Units Sold", number(float(kpis.get("totalUnits", 0) or 0))),
            ("Sold SKUs", number(float(kpis.get("soldSkus", 0) or 0))),
            ("Brands", number(float(kpis.get("brands", 0) or 0))),
            ("Customers", number(float(kpis.get("customers", 0) or 0))),
        ],
        "kpi",
    )

    brand_tab, customer_tab, product_tab, raw_tab = st.tabs(
        ["Brand Dashboard", "Customer Dashboard", "Product Dashboard", "Raw Data"]
    )

    with brand_tab:
        st.markdown('<div class="section-heading"><h3>Brand Performance</h3></div>', unsafe_allow_html=True)
        chart_left, chart_right = st.columns(2)
        with chart_left:
            render_bar_panel("Top Brand Sales", brand_df, "Brand", "Revenue", money)
        with chart_right:
            render_bar_panel("Top Brand Units Sold", brand_df, "Brand", "Quantity", number, "blue")

        brand_query = st.text_input(
            "Search brand, product, SKU, barcode",
            key=f"{key_prefix}_brand_query",
        )
        filtered = brand_df
        if brand_query and not filtered.empty:
            filtered = filtered[filtered["Brand"].str.contains(brand_query, case=False, na=False)]
        filtered_display = filtered.copy()
        if "Share" in filtered_display:
            filtered_display["Share"] = filtered_display["Share"].astype(float) * 100
        st.dataframe(
            filtered_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Revenue": st.column_config.NumberColumn(format="$%.2f"),
                "Share": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )

    with customer_tab:
        st.markdown('<div class="section-heading"><h3>Customer Performance</h3></div>', unsafe_allow_html=True)
        chart_left, chart_right = st.columns(2)
        with chart_left:
            render_bar_panel("Top Customer Sales", customer_df, "Customer Name", "Revenue", money)
        with chart_right:
            render_bar_panel("Top Customer Units Sold", customer_df, "Customer Name", "Quantity", number, "blue")

        customer_query = st.text_input(
            "Search customer, brand, product, SKU, barcode",
            key=f"{key_prefix}_customer_query",
        )
        filtered = customer_df
        if customer_query and not filtered.empty:
            summary_mask = filtered["Customer Name"].str.contains(customer_query, case=False, na=False)
            detail_matches = pd.Series(False, index=customer_detail_df.index)
            if not customer_detail_df.empty:
                for column in ["Customer Name", "Brand", "Product Name", "SKU", "Barcode Text"]:
                    detail_matches = detail_matches | customer_detail_df[column].astype(str).str.contains(
                        customer_query,
                        case=False,
                        na=False,
                    )
            matching_customers = set(customer_detail_df.loc[detail_matches, "Customer Name"]) if not customer_detail_df.empty else set()
            filtered = filtered[summary_mask | filtered["Customer Name"].isin(matching_customers)]
        filtered_display = filtered.copy()
        if "Share" in filtered_display:
            filtered_display["Share"] = filtered_display["Share"].astype(float) * 100
        st.dataframe(
            filtered_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Revenue": st.column_config.NumberColumn(format="$%.2f"),
                "Share": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )

    with product_tab:
        st.markdown('<div class="section-heading"><h3>Product Performance</h3></div>', unsafe_allow_html=True)
        product_query = st.text_input(
            "Search brand, product, SKU, barcode",
            key=f"{key_prefix}_product_query",
        )
        filtered = product_df
        if product_query and not filtered.empty:
            search_cols = ["Brand", "Product Name", "SKU", "Barcode Text"]
            mask = pd.Series(False, index=filtered.index)
            for column in search_cols:
                mask = mask | filtered[column].astype(str).str.contains(product_query, case=False, na=False)
            filtered = filtered[mask]
        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Revenue": st.column_config.NumberColumn(format="$%.2f"),
                "MinPrice": st.column_config.NumberColumn(format="$%.2f"),
                "MaxPrice": st.column_config.NumberColumn(format="$%.2f"),
            },
        )

        if not product_monthly_df.empty:
            st.markdown('<div class="section-heading"><h3>Monthly Product Details</h3></div>', unsafe_allow_html=True)
            month_options = sorted(product_monthly_df["Order Month"].dropna().unique())
            selected_month = st.selectbox("Month", month_options, key=f"{key_prefix}_month")
            month_rows = product_monthly_df[product_monthly_df["Order Month"].eq(selected_month)]
            month_rows = month_rows.sort_values("Quantity", ascending=False).head(100)
            st.dataframe(
                month_rows,
                use_container_width=True,
                hide_index=True,
                column_config={"Revenue": st.column_config.NumberColumn(format="$%.2f")},
            )

        if not product_matrix_df.empty:
            st.markdown('<div class="section-heading"><h3>Product Monthly Units Matrix</h3></div>', unsafe_allow_html=True)
            st.dataframe(product_matrix_df, use_container_width=True, hide_index=True)

    with raw_tab:
        st.subheader("Archive Metadata")
        st.json(metadata)
        st.download_button(
            "Download dashboard data",
            gzip.compress(json.dumps(report_data, ensure_ascii=False).encode("utf-8")),
            file_name=f"{safe_filename(metadata.get('reportTitle', 'dashboard'))}.json.gz",
            mime="application/gzip",
            key=f"{key_prefix}_download_data",
        )


def generate_dashboard(sales_files, products_file, report_title: str) -> None:
    if not GENERATOR.exists():
        st.error(f"Generator file not found: {GENERATOR}")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        try:
            sales_path = combine_sales_files(sales_files, tmp_path / "combined_sales_orders.xlsx")
        except Exception as exc:
            st.error("Sales Orders files could not be merged.")
            st.code(str(exc))
            return

        products_path = save_uploaded_file(products_file, tmp_path)

        filename = safe_filename(report_title)
        html_output = tmp_path / f"{filename}.html"
        excel_output = tmp_path / f"{filename}.xlsx"
        data_output = tmp_path / f"{filename}.json.gz"

        result = subprocess.run(
            [
                sys.executable,
                str(GENERATOR),
                "--sales-file",
                str(sales_path),
                "--products-file",
                str(products_path),
                "--html-output",
                str(html_output),
                "--excel-output",
                str(excel_output),
                "--data-output",
                str(data_output),
                "--report-title",
                report_title,
            ],
            capture_output=True,
            text=True,
        )

        if result.returncode != 0:
            st.error("Report generation failed.")
            st.code(result.stderr or result.stdout)
            return

        archive_base = f"{datetime.now():%Y-%m-%d-%H%M%S}-{filename}"
        data_archive_path = REPORT_DIR / f"{archive_base}.json.gz"
        data_archive_path.write_bytes(data_output.read_bytes())
        load_dashboard_data.clear()
        st.success(f"Dashboard saved. Open Saved Dashboards to view it: {archive_base}")

        st.download_button(
            "Download Excel Report",
            excel_output.read_bytes(),
            file_name=f"{filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Download HTML Report",
            html_output.read_text(encoding="utf-8"),
            file_name=f"{filename}.html",
            mime="text/html",
        )


st.set_page_config(page_title="Nakama Sales Dashboard", layout="wide")
st.markdown(
    """
    <style>
    :root {
        --nakama-ink: #172026;
        --nakama-muted: #65717d;
        --nakama-line: #d9dee4;
        --nakama-panel: #f7f8fa;
        --nakama-accent: #ff4b4b;
        --nakama-accent-2: #2f6f9f;
        --nakama-accent-soft: #ffe1e1;
        --nakama-accent-2-soft: #d9eaf2;
    }

    html,
    body,
    .stApp,
    .block-container {
        color: var(--nakama-ink);
        font-family: Arial, "Microsoft YaHei", sans-serif;
        font-size: 15px;
    }

    body,
    [data-testid="stAppViewContainer"] {
        background: #ffffff;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    .block-container {
        max-width: 1800px;
        padding: 32px 40px 42px;
    }

    h1 {
        margin: 0 0 8px !important;
        color: var(--nakama-ink);
        font-size: 30px !important;
        line-height: 1.2 !important;
        letter-spacing: 0 !important;
    }

    h2 {
        margin: 0 0 14px !important;
        color: var(--nakama-ink);
        font-size: 18px !important;
        letter-spacing: 0 !important;
    }

    h3 {
        margin: 0 0 10px !important;
        color: var(--nakama-ink);
        font-size: 15px !important;
        letter-spacing: 0 !important;
    }

    p,
    [data-testid="stCaptionContainer"] {
        color: var(--nakama-muted);
    }

    [data-testid="stTabs"] [role="tablist"] {
        gap: 8px;
        margin: 12px 0 22px;
        border-bottom: 1px solid var(--nakama-line);
    }

    [data-testid="stTabs"] [role="tab"] {
        height: auto;
        margin-bottom: 0;
        padding: 10px 14px;
        border: 1px solid var(--nakama-line) !important;
        border-bottom: 0 !important;
        border-radius: 8px 8px 0 0 !important;
        background: #ffffff !important;
        color: var(--nakama-ink) !important;
        font-weight: 700;
    }

    [data-testid="stTabs"] [role="tab"][aria-selected="true"] {
        background: var(--nakama-accent) !important;
        border-color: var(--nakama-accent) !important;
        color: #ffffff !important;
    }

    [data-testid="stTabs"] [role="tab"] p {
        color: inherit;
        font-weight: inherit;
    }

    .report-header {
        margin: 6px 0 14px;
    }

    .report-header h2 {
        margin: 0 0 6px !important;
        font-size: 22px !important;
    }

    .report-header p {
        margin: 0;
        color: var(--nakama-muted);
    }

    .highlights {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
        margin-bottom: 12px;
    }

    .highlight {
        min-height: 88px;
        padding: 16px;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
        background: #ffffff;
    }

    .highlight .value {
        font-size: 18px;
        line-height: 1.25;
        overflow-wrap: anywhere;
    }

    .kpis {
        display: grid;
        grid-template-columns: repeat(5, minmax(140px, 1fr));
        gap: 12px;
        margin-bottom: 24px;
    }

    .kpi {
        min-height: 110px;
        padding: 14px;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
        background: var(--nakama-panel);
    }

    .label {
        color: var(--nakama-muted);
        font-size: 12px;
        font-weight: 700;
        letter-spacing: .05em;
        text-transform: uppercase;
    }

    .value {
        margin-top: 5px;
        color: var(--nakama-ink);
        font-size: 22px;
        font-weight: 700;
    }

    .section-heading {
        display: flex;
        align-items: end;
        justify-content: space-between;
        gap: 16px;
        margin: 18px 0 10px;
    }

    .chart-card {
        min-height: 260px;
        padding: 14px;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
        background: #ffffff;
        margin: 8px 0 18px;
    }

    .chart-card h3 {
        margin: 0 0 10px !important;
    }

    .bar-row {
        display: grid;
        grid-template-columns: minmax(160px, 260px) 1fr minmax(90px, auto);
        gap: 10px;
        align-items: center;
        margin: 8px 0;
    }

    .bar-label {
        font-weight: 700;
        overflow-wrap: anywhere;
    }

    .bar-track {
        height: 18px;
        overflow: hidden;
        border-radius: 5px;
        background: var(--nakama-accent-soft);
    }

    .bar-fill {
        height: 100%;
        background: var(--nakama-accent);
    }

    .bar-fill.blue {
        background: var(--nakama-accent-2);
    }

    .bar-value {
        text-align: right;
        font-weight: 700;
    }

    [data-testid="stAlert"] {
        border: 1px solid var(--nakama-line) !important;
        border-radius: 8px !important;
        background: #ffffff !important;
        color: var(--nakama-ink);
    }

    [data-testid="stTextInput"] input,
    [data-baseweb="select"] > div,
    [data-baseweb="input"] > div {
        border-color: var(--nakama-line) !important;
        border-radius: 6px !important;
        background: #ffffff !important;
        color: var(--nakama-ink) !important;
    }

    div[data-testid="stDownloadButton"] button {
        min-height: 42px;
        padding: 10px 14px;
        border-radius: 6px !important;
        border-color: var(--nakama-line) !important;
        font-weight: 700;
    }

    div[data-testid="stButton"] button[kind="primary"] {
        border-color: var(--nakama-accent) !important;
        background: var(--nakama-accent) !important;
        color: #ffffff !important;
    }

    div[data-testid="stDownloadButton"] button {
        background: #ffffff !important;
        color: var(--nakama-accent-2) !important;
    }

    [data-testid="stDataFrame"],
    [data-testid="stTable"] {
        overflow: hidden;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
    }

    [data-testid="stJson"] {
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
        background: var(--nakama-panel);
    }

    @media (max-width: 900px) {
        .block-container {
            padding-left: 18px;
            padding-right: 18px;
        }

        .highlights,
        .kpis {
            grid-template-columns: 1fr;
        }

        .bar-row {
            grid-template-columns: 1fr;
            gap: 5px;
        }

        .bar-value {
            text-align: left;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("Nakama Sales Dashboard")

generate_tab, saved_tab = st.tabs(["Generate Dashboard", "Saved Dashboards"])

if "product_upload_key" not in st.session_state:
    st.session_state.product_upload_key = 0

with generate_tab:
    st.subheader("Generate New Dashboard")
    sales_files = st.file_uploader(
        "Upload Sales Orders Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

    st.write("Upload Products Excel")
    products_file = st.session_state.get("products_file")
    if products_file:
        product_name_col, product_action_col = st.columns([4, 1])
        product_name_col.info(f"Selected product file: {products_file['name']}")
        if product_action_col.button("Change Products Excel", key="change_products_file"):
            st.session_state.pop("products_file", None)
            st.session_state.product_upload_key += 1
            st.rerun()
    else:
        uploaded_products_file = st.file_uploader(
            "Upload Products Excel",
            type=["xlsx", "xls"],
            key=f"products_file_{st.session_state.product_upload_key}",
            label_visibility="collapsed",
        )
        if uploaded_products_file:
            st.session_state.products_file = uploaded_file_state(uploaded_products_file)
            st.rerun()

    report_title = st.text_input("Report Title", value="Nakama Sales Report")

    if sales_files and products_file:
        if st.button("Generate Dashboard", type="primary"):
            generate_dashboard(sales_files, products_file, report_title)
    else:
        st.info("Upload at least one Sales Orders Excel file and one Products Excel file to generate a dashboard.")

with saved_tab:
    st.subheader("Saved Dashboards")
    saved_reports = saved_dashboard_files()

    if not saved_reports:
        st.info("No saved dashboards found. Add .json.gz reports to the reports folder.")
    else:
        selected = st.selectbox(
            "Select a saved dashboard",
            saved_reports,
            format_func=saved_dashboard_label,
        )

        download_col, delete_col = st.columns([1, 1])
        with download_col:
            st.download_button(
                "Download dashboard file",
                selected.read_bytes(),
                file_name=selected.name,
                mime="application/gzip" if selected.name.endswith(".json.gz") else "text/html",
            )
        with delete_col:
            confirm_delete = st.checkbox(
                "Confirm delete",
                key=f"confirm_delete_{selected.name}",
            )
            if st.button(
                "Delete selected dashboard",
                disabled=not confirm_delete,
                type="secondary",
                key=f"delete_{selected.name}",
            ):
                delete_saved_dashboard(selected)

        data = load_dashboard_data(str(selected))
        render_dashboard(data, "saved")
