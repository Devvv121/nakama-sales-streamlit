from __future__ import annotations

import gzip
import json
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


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
    path = folder / uploaded_file.name
    path.write_bytes(uploaded_file.getbuffer())
    return path


def saved_dashboard_label(path: Path) -> str:
    if path.name.endswith(".json.gz"):
        return path.name.removesuffix(".json.gz")
    return path.stem


def delete_saved_dashboard(path: Path) -> None:
    path.unlink()
    load_dashboard_data.clear()
    st.success(f"Deleted dashboard: {saved_dashboard_label(path)}")
    st.rerun()


def render_dashboard(report_data: dict, key_prefix: str) -> None:
    metadata = report_data.get("metadata", {})
    kpis = report_data.get("kpis", {})

    st.caption(kpis.get("dateRange", ""))

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total Sales", money(float(kpis.get("totalSales", 0))))
    col2.metric("Units Sold", number(float(kpis.get("totalUnits", 0))))
    col3.metric("Sold SKUs", number(float(kpis.get("soldSkus", 0))))
    col4.metric("Brands", number(float(kpis.get("brands", 0))))
    col5.metric("Customers", number(float(kpis.get("customers", 0))))

    top1, top2, top3 = st.columns(3)
    top1.info(f"Top Brand: {kpis.get('topBrand', '')}")
    top2.info(f"Top Customer: {kpis.get('topCustomer', '')}")
    top3.info(f"Top 10 Brand Share: {percent(float(kpis.get('top10BrandShare', 0)))}")

    brand_df = pd.DataFrame(report_data.get("brands", []))
    customer_df = pd.DataFrame(report_data.get("customers", []))
    product_df = pd.DataFrame(report_data.get("products", []))
    product_monthly_df = pd.DataFrame(report_data.get("productMonthlyDetails", []))
    product_matrix_df = pd.DataFrame(report_data.get("productMonthlyMatrix", []))

    dashboard_tab, brand_tab, customer_tab, product_tab, matrix_tab, raw_tab = st.tabs(
        ["Overview", "Brands", "Customers", "Products", "Monthly Matrix", "Raw Data"]
    )

    with dashboard_tab:
        left, right = st.columns(2)
        if not brand_df.empty:
            left.subheader("Top Brands by Sales")
            brand_chart = brand_df.head(15).set_index("Brand")["Revenue"]
            left.bar_chart(brand_chart)
        if not customer_df.empty:
            right.subheader("Top Customers by Sales")
            customer_chart = customer_df.head(15).set_index("Customer Name")["Revenue"]
            right.bar_chart(customer_chart)

        if not product_df.empty:
            st.subheader("Top Products by Units")
            top_products = product_df.sort_values("Quantity", ascending=False).head(50)
            st.dataframe(top_products, use_container_width=True, hide_index=True)

    with brand_tab:
        st.subheader("Brand Summary")
        brand_query = st.text_input("Filter brand", key=f"{key_prefix}_brand_query")
        filtered = brand_df
        if brand_query and not filtered.empty:
            filtered = filtered[filtered["Brand"].str.contains(brand_query, case=False, na=False)]
        st.dataframe(filtered, use_container_width=True, hide_index=True)

    with customer_tab:
        st.subheader("Customer Summary")
        customer_query = st.text_input("Filter customer", key=f"{key_prefix}_customer_query")
        filtered = customer_df
        if customer_query and not filtered.empty:
            filtered = filtered[filtered["Customer Name"].str.contains(customer_query, case=False, na=False)]
        st.dataframe(filtered, use_container_width=True, hide_index=True)

    with product_tab:
        st.subheader("Product Details")
        product_query = st.text_input("Filter product, SKU, barcode, or brand", key=f"{key_prefix}_product_query")
        filtered = product_df
        if product_query and not filtered.empty:
            search_cols = ["Brand", "Product Name", "SKU", "Barcode Text"]
            mask = pd.Series(False, index=filtered.index)
            for column in search_cols:
                mask = mask | filtered[column].astype(str).str.contains(product_query, case=False, na=False)
            filtered = filtered[mask]
        st.dataframe(filtered, use_container_width=True, hide_index=True)

        if not product_monthly_df.empty:
            st.subheader("Monthly Product Details")
            month_options = sorted(product_monthly_df["Order Month"].dropna().unique())
            selected_month = st.selectbox("Month", month_options, key=f"{key_prefix}_month")
            month_rows = product_monthly_df[product_monthly_df["Order Month"].eq(selected_month)]
            month_rows = month_rows.sort_values("Quantity", ascending=False).head(100)
            st.dataframe(month_rows, use_container_width=True, hide_index=True)

    with matrix_tab:
        st.subheader("Product Monthly Units Matrix")
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


def generate_dashboard(sales_file, products_file, report_title: str) -> None:
    if not GENERATOR.exists():
        st.error(f"Generator file not found: {GENERATOR}")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        sales_path = save_uploaded_file(sales_file, tmp_path)
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

        archive_name = f"{datetime.now():%Y-%m-%d-%H%M%S}-{filename}.json.gz"
        archive_path = REPORT_DIR / archive_name
        archive_path.write_bytes(data_output.read_bytes())
        load_dashboard_data.clear()
        st.success(f"Dashboard saved. Open Saved Dashboards to view it: {archive_name}")

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
st.title("Nakama Sales Dashboard")

generate_tab, saved_tab = st.tabs(["Generate Dashboard", "Saved Dashboards"])

with generate_tab:
    st.subheader("Generate New Dashboard")
    sales_file = st.file_uploader("Upload Sales Orders Excel", type=["xlsx", "xls"])
    products_file = st.file_uploader("Upload Products Excel", type=["xlsx", "xls"])
    report_title = st.text_input("Report Title", value="Nakama Sales Report")

    if sales_file and products_file:
        if st.button("Generate Dashboard", type="primary"):
            generate_dashboard(sales_file, products_file, report_title)
    else:
        st.info("Upload both Excel files to generate a dashboard.")

with saved_tab:
    st.subheader("Saved Dashboards")
    saved_reports = sorted(
        [*REPORT_DIR.glob("*.json.gz"), *REPORT_DIR.glob("*.html")],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

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

        if selected.name.endswith(".json.gz"):
            data = load_dashboard_data(str(selected))
            render_dashboard(data, "saved")
        else:
            html = selected.read_text(encoding="utf-8")
            st.warning("This is an old HTML archive. Convert future reports to .json.gz for faster loading.")
            components.html(html, height=1800, scrolling=True)
