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
import streamlit.components.v1 as components

try:
    from st_keyup import st_keyup
except ImportError:
    st_keyup = None


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


def live_search_input(label: str, key: str) -> str:
    st.markdown(f'<div class="search-label">{escape(label)}</div>', unsafe_allow_html=True)
    st.markdown('<div class="search-box-wrap">', unsafe_allow_html=True)
    if st_keyup is None:
        value = st.text_input(" ", key=key, placeholder="Type to search...", label_visibility="collapsed")
    else:
        value = st_keyup(
            " ",
            key=key,
            placeholder="Type to search...",
            label_visibility="collapsed",
        ) or ""
    st.markdown("</div>", unsafe_allow_html=True)
    return value


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


def row_search_mask(frame: pd.DataFrame, query: str, columns: list[str]) -> pd.Series:
    mask = pd.Series(False, index=frame.index)
    for column in columns:
        if column in frame:
            mask = mask | frame[column].astype(str).str.contains(query, case=False, na=False)
    return mask


def display_money_columns(columns: list[str]) -> dict:
    return {column: st.column_config.NumberColumn(format="$%.2f") for column in columns}


def format_table_value(value, column: str) -> str:
    if pd.isna(value):
        return ""
    if column in {"Revenue", "MinPrice", "MaxPrice"}:
        return money(float(value or 0))
    if column in {"Share"}:
        return f"{float(value or 0):.1f}%"
    if column in {"Quantity", "SkuCount", "Orders"}:
        return number(float(value or 0))
    return str(value)


def render_html_table(frame: pd.DataFrame, max_rows: int = 300) -> None:
    if frame.empty:
        st.info("No matching rows.")
        return

    display = frame.head(max_rows).copy()
    header = "".join(f"<th>{escape(str(column))}</th>" for column in display.columns)
    rows = []
    for _, row in display.iterrows():
        cells = "".join(
            f"<td>{escape(format_table_value(row[column], str(column)))}</td>"
            for column in display.columns
        )
        rows.append(f"<tr>{cells}</tr>")

    st.markdown(
        '<div class="html-table-box">'
        '<table class="html-data-table">'
        f"<thead><tr>{header}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
        "</div>",
        unsafe_allow_html=True,
    )
    if len(frame) > max_rows:
        st.info(f"Showing the first {max_rows} rows. Use search to narrow the table.")


def ensure_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for column in columns:
        if column not in frame:
            frame[column] = pd.Series(dtype="object")
    return frame


def product_monthly_rows(product_monthly_df: pd.DataFrame, product: pd.Series) -> pd.DataFrame:
    if product_monthly_df.empty:
        return product_monthly_df

    rows = product_monthly_df
    for column in ["Brand", "Product Name", "SKU"]:
        if column in rows and column in product:
            rows = rows[rows[column].astype(str).eq(str(product[column]))]
    return rows.sort_values("Order Month")


def render_product_expanders(
    products: pd.DataFrame,
    product_monthly_df: pd.DataFrame,
    key_prefix: str,
    limit: int = 40,
) -> None:
    if products.empty:
        st.info("No matching products.")
        return

    for idx, (_, product) in enumerate(products.head(limit).iterrows()):
        summary = (
            f"{product.get('Product Name', '')} | Barcode: {product.get('Barcode Text', '')} | "
            f"{product.get('Brand', '')} | "
            f"Units: {number(float(product.get('Quantity', 0) or 0))} | "
            f"Sales: {money(float(product.get('Revenue', 0) or 0))}"
        )
        with st.expander(summary):
            render_html_table(pd.DataFrame([product]), max_rows=1)
            monthly = product_monthly_rows(product_monthly_df, product)
            if monthly.empty:
                st.info("No monthly details for this product.")
            else:
                render_html_table(monthly[["Order Month", "Quantity", "Revenue"]], max_rows=60)

    if len(products) > limit:
        st.info(f"Showing the first {limit} matching products. Use search to narrow the list.")


def render_dashboard(report_data: dict, key_prefix: str) -> None:
    metadata = report_data.get("metadata", {})
    kpis = report_data.get("kpis", {})

    brand_df = pd.DataFrame(report_data.get("brands", []))
    customer_df = pd.DataFrame(report_data.get("customers", []))
    product_df = pd.DataFrame(report_data.get("products", []))
    customer_detail_df = pd.DataFrame(report_data.get("customerDetails", []))
    product_monthly_df = pd.DataFrame(report_data.get("productMonthlyDetails", []))
    product_matrix_df = pd.DataFrame(report_data.get("productMonthlyMatrix", []))
    brand_df = ensure_columns(brand_df, ["Brand", "Quantity", "Revenue", "SkuCount", "Share"])
    customer_df = ensure_columns(customer_df, ["Customer Name", "Revenue", "Quantity", "Orders", "SkuCount", "Share"])
    product_df = ensure_columns(product_df, ["Brand", "Product Name", "SKU", "Barcode Text", "Quantity", "Revenue", "MinPrice", "MaxPrice"])
    customer_detail_df = ensure_columns(customer_detail_df, ["Customer Name", "Brand", "Product Name", "SKU", "Barcode Text", "Quantity", "Revenue", "MinPrice", "MaxPrice"])
    product_monthly_df = ensure_columns(product_monthly_df, ["Brand", "Product Name", "SKU", "Barcode Text", "Order Month", "Quantity", "Revenue"])

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

        brand_query = live_search_input(
            "Search brand, product, SKU, barcode",
            key=f"{key_prefix}_brand_query",
        )
        filtered = brand_df
        if brand_query and not filtered.empty:
            brand_mask = filtered["Brand"].str.contains(brand_query, case=False, na=False)
            product_matches = row_search_mask(
                product_df,
                brand_query,
                ["Brand", "Product Name", "SKU", "Barcode Text"],
            )
            matching_brands = set(product_df.loc[product_matches, "Brand"]) if not product_df.empty else set()
            filtered = filtered[brand_mask | filtered["Brand"].isin(matching_brands)]
        filtered_display = filtered.copy()
        if "Share" in filtered_display:
            filtered_display["Share"] = filtered_display["Share"].astype(float) * 100
        render_html_table(filtered_display, max_rows=300)
        st.markdown('<div class="section-heading"><h3>Brand Details</h3></div>', unsafe_allow_html=True)
        if not brand_query:
            st.info("Search a brand, product, SKU, or barcode to show expandable brand details.")
        else:
            detail_limit = 40
            for idx, (_, brand) in enumerate(filtered.head(detail_limit).iterrows()):
                brand_name = str(brand.get("Brand", ""))
                brand_products = product_df[product_df["Brand"].astype(str).eq(brand_name)]
                product_mask = row_search_mask(
                    brand_products,
                    brand_query,
                    ["Brand", "Product Name", "SKU", "Barcode Text"],
                )
                if brand_query.lower() not in brand_name.lower():
                    brand_products = brand_products[product_mask]
                with st.expander(
                    f"{brand_name} | Units: {number(float(brand.get('Quantity', 0) or 0))} | "
                    f"Sales: {money(float(brand.get('Revenue', 0) or 0))} | SKUs: {number(float(brand.get('SkuCount', 0) or 0))}"
                ):
                    render_product_expanders(
                        brand_products.sort_values("Quantity", ascending=False),
                        product_monthly_df,
                        f"{key_prefix}_brand_{idx}",
                        limit=40,
                    )
            if len(filtered) > detail_limit:
                st.info(f"Showing the first {detail_limit} matching brands. Use search to narrow the list.")

    with customer_tab:
        st.markdown('<div class="section-heading"><h3>Customer Performance</h3></div>', unsafe_allow_html=True)
        chart_left, chart_right = st.columns(2)
        with chart_left:
            render_bar_panel("Top Customer Sales", customer_df, "Customer Name", "Revenue", money)
        with chart_right:
            render_bar_panel("Top Customer Units Sold", customer_df, "Customer Name", "Quantity", number, "blue")

        customer_query = live_search_input(
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
        render_html_table(filtered_display, max_rows=300)
        st.markdown('<div class="section-heading"><h3>Customer Details</h3></div>', unsafe_allow_html=True)
        if not customer_query:
            st.info("Search a customer, brand, product, SKU, or barcode to show expandable customer details.")
        else:
            detail_limit = 40
            for idx, (_, customer) in enumerate(filtered.head(detail_limit).iterrows()):
                customer_name = str(customer.get("Customer Name", ""))
                details = customer_detail_df[customer_detail_df["Customer Name"].astype(str).eq(customer_name)]
                if not details.empty:
                    detail_mask = row_search_mask(
                        details,
                        customer_query,
                        ["Customer Name", "Brand", "Product Name", "SKU", "Barcode Text"],
                    )
                    if customer_query.lower() not in customer_name.lower():
                        details = details[detail_mask]
                with st.expander(
                    f"{customer_name} | Sales: {money(float(customer.get('Revenue', 0) or 0))} | "
                    f"Units: {number(float(customer.get('Quantity', 0) or 0))} | "
                    f"Orders: {number(float(customer.get('Orders', 0) or 0))}"
                ):
                    render_html_table(details.sort_values("Quantity", ascending=False), max_rows=300)
            if len(filtered) > detail_limit:
                st.info(f"Showing the first {detail_limit} matching customers. Use search to narrow the list.")

    with product_tab:
        st.markdown('<div class="section-heading"><h3>Product Performance</h3></div>', unsafe_allow_html=True)
        product_query = live_search_input(
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
        render_html_table(filtered, max_rows=300)
        st.markdown('<div class="section-heading"><h3>Expandable Product Monthly Details</h3></div>', unsafe_allow_html=True)
        if not product_query:
            st.info("Search a product, brand, SKU, or barcode to show expandable monthly details.")
        else:
            render_product_expanders(
                filtered.sort_values("Quantity", ascending=False),
                product_monthly_df,
                f"{key_prefix}_product",
                limit=40,
            )

        if not product_monthly_df.empty:
            st.markdown('<div class="section-heading"><h3>Monthly Product Details</h3></div>', unsafe_allow_html=True)
            month_options = sorted(product_monthly_df["Order Month"].dropna().unique())
            selected_month = st.selectbox("Month", month_options, key=f"{key_prefix}_month")
            month_rows = product_monthly_df[product_monthly_df["Order Month"].eq(selected_month)]
            month_rows = month_rows.sort_values("Quantity", ascending=False).head(100)
            render_html_table(month_rows, max_rows=300)

        if not product_matrix_df.empty:
            st.markdown('<div class="section-heading"><h3>Product Monthly Units Matrix</h3></div>', unsafe_allow_html=True)
            render_html_table(product_matrix_df, max_rows=300)

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
        --nakama-muted: #172026;
        --nakama-line: #000000;
        --nakama-panel: #ffffff;
        --nakama-accent: #ff4b4b;
        --nakama-accent-2: #2f6f9f;
        --nakama-accent-soft: #ffffff;
        --nakama-accent-2-soft: #ffffff;
    }

    html,
    body,
    .stApp,
    .block-container {
        color: var(--nakama-ink);
        font-family: Arial, "Microsoft YaHei", sans-serif;
        font-size: 18px;
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
        font-size: 33px !important;
        line-height: 1.2 !important;
        letter-spacing: 0 !important;
    }

    h2 {
        margin: 0 0 14px !important;
        color: var(--nakama-ink);
        font-size: 21px !important;
        letter-spacing: 0 !important;
    }

    h3 {
        margin: 0 0 10px !important;
        color: var(--nakama-ink);
        font-size: 18px !important;
        letter-spacing: 0 !important;
    }

    p,
    label,
    span,
    small,
    div,
    [data-testid="stCaptionContainer"],
    [data-testid="stMarkdownContainer"],
    [data-testid="stWidgetLabel"],
    [data-testid="stFileUploader"] *,
    [data-testid="stAlert"] *,
    [data-testid="stExpander"] *,
    [data-testid="stDataFrame"] *,
    [data-testid="stTable"] * {
        color: var(--nakama-ink) !important;
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
        color: var(--nakama-ink) !important;
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
        font-size: 24px !important;
    }

    .report-header p {
        margin: 0;
        color: var(--nakama-ink);
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
        font-size: 20px;
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
        color: var(--nakama-ink);
        font-size: 14px;
        font-weight: 700;
        letter-spacing: .05em;
        text-transform: uppercase;
    }

    .value {
        margin-top: 5px;
        color: var(--nakama-ink);
        font-size: 24px;
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

    .search-label {
        margin: 14px 0 8px;
        color: var(--nakama-ink) !important;
        font-size: 28px;
        font-weight: 900;
        line-height: 1.25;
    }

    .search-box-wrap {
        margin: 0 0 18px;
    }

    .search-box-wrap input,
    .search-box-wrap [data-baseweb="input"] > div,
    .search-box-wrap [data-baseweb="input"] input {
        min-height: 46px !important;
        border: 2px solid var(--nakama-line) !important;
        border-radius: 8px !important;
        background: #ffffff !important;
        color: var(--nakama-ink) !important;
        font-size: 18px !important;
        padding: 8px 12px !important;
        box-shadow: none !important;
        outline: none !important;
    }

    .search-box-wrap input::placeholder {
        color: var(--nakama-ink) !important;
        opacity: 0.7 !important;
    }

    .search-box-wrap input:focus,
    .search-box-wrap [data-baseweb="input"] > div:focus-within {
        border-color: var(--nakama-line) !important;
        box-shadow: 0 0 0 2px var(--nakama-line) !important;
    }

    [data-testid="stFileUploader"] {
        border: 1px solid var(--nakama-line) !important;
        border-radius: 8px !important;
        background: #ffffff !important;
        padding: 12px 14px !important;
        margin: 8px 0 18px !important;
    }

    [data-testid="stFileUploader"] section,
    [data-testid="stFileUploaderDropzone"] {
        min-height: 82px !important;
        border: 1px solid var(--nakama-line) !important;
        border-radius: 8px !important;
        background: #ffffff !important;
        color: var(--nakama-ink) !important;
    }

    [data-testid="stExpander"] {
        border: 1px solid var(--nakama-line) !important;
        border-radius: 8px !important;
        background: #ffffff !important;
        margin: 8px 0 !important;
    }

    [data-testid="stExpander"] summary {
        color: var(--nakama-ink) !important;
        font-weight: 700;
        min-height: 38px !important;
        padding: 8px 12px !important;
    }

    [data-testid="stExpander"] details > div {
        padding: 8px 12px !important;
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
        color: var(--nakama-ink) !important;
    }

    div[data-testid="stDownloadButton"] button {
        background: #ffffff !important;
        color: var(--nakama-ink) !important;
    }

    [data-testid="stDataFrame"],
    [data-testid="stTable"] {
        overflow: hidden;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
    }

    [data-testid="stDataFrame"] [role="columnheader"],
    [data-testid="stDataFrame"] [role="gridcell"],
    [data-testid="stDataFrame"] [role="rowheader"],
    [data-testid="stDataFrame"] [data-testid*="column"],
    [data-testid="stDataFrame"] [class*="header"],
    [data-testid="stDataFrame"] [class*="Header"],
    [data-testid="stTable"] th,
    [data-testid="stTable"] td {
        color: var(--nakama-ink) !important;
        -webkit-text-fill-color: var(--nakama-ink) !important;
        opacity: 1 !important;
    }

    [data-testid="stDataFrame"] canvas {
        opacity: 1 !important;
        filter: contrast(1.25) saturate(0) !important;
    }

    .html-table-box {
        width: 100%;
        max-height: 460px;
        overflow: auto;
        padding: 0;
        border: 1px solid var(--nakama-line);
        border-radius: 8px;
        margin: 6px 0 16px;
        background: #ffffff;
    }

    .html-data-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 17px;
        color: var(--nakama-ink);
        line-height: 1.2;
    }

    .html-data-table tbody tr {
        height: 34px;
    }

    .html-data-table th,
    .html-data-table td {
        padding: 6px 10px;
        border-bottom: 1px solid var(--nakama-line);
        color: var(--nakama-ink) !important;
        text-align: left;
        vertical-align: top;
        white-space: nowrap;
        line-height: 1.2;
    }

    .html-data-table th {
        position: sticky;
        top: 0;
        z-index: 1;
        background: #ffffff;
        border-bottom: 2px solid var(--nakama-line);
        font-weight: 700;
    }

    .html-data-table tr:last-child td {
        border-bottom: 1px solid var(--nakama-line);
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
components.html(
    """
    <script>
    (() => {
      const doc = window.parent.document;
      if (doc.__nakamaShortcutGuardInstalled) return;
      doc.__nakamaShortcutGuardInstalled = true;

      const isEditable = (target) => {
        if (!target) return false;
        const tag = (target.tagName || "").toLowerCase();
        return tag === "input" || tag === "textarea" || target.isContentEditable;
      };

      const guardShortcut = (event) => {
          const key = (event.key || "").toLowerCase();
          if ((event.ctrlKey || event.metaKey) && (key === "c" || key === "v") && !isEditable(event.target)) {
            event.stopImmediatePropagation();
          }
      };

      ["keydown", "keyup"].forEach((eventName) => {
        window.parent.addEventListener(eventName, guardShortcut, true);
        doc.addEventListener(eventName, guardShortcut, true);
      });
    })();
    </script>
    """,
    height=0,
)

saved_tab, generate_tab = st.tabs(["Saved Dashboards", "Generate Dashboard"])

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
