import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Nakama Sales Dashboard", layout="wide")

st.title("Nakama Sales Dashboard")

generate_tab, archive_tab = st.tabs(["Generate Dashboard", "Saved Dashboards"])

REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(exist_ok=True)
GENERATOR = Path(__file__).parent / "Generate-SalesBrandReport.py"


with generate_tab:
    st.subheader("Generate New Dashboard")

    sales_file = st.file_uploader("Upload Sales Orders Excel", type=["xlsx", "xls"])
    products_file = st.file_uploader("Upload Products Excel", type=["xlsx", "xls"])
    report_title = st.text_input("Report Title", value="Nakama Sales Report")

    save_to_archive = st.checkbox("Save this dashboard to history", value=True)

    if sales_file and products_file:
        if st.button("Generate Dashboard", type="primary"):
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)

                sales_path = tmp_path / sales_file.name
                products_path = tmp_path / products_file.name
                html_output = tmp_path / f"{report_title}.html"
                excel_output = tmp_path / f"{report_title}.xlsx"

                sales_path.write_bytes(sales_file.getbuffer())
                products_path.write_bytes(products_file.getbuffer())

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
                        "--report-title",
                        report_title,
                    ],
                    capture_output=True,
                    text=True,
                )

                if result.returncode != 0:
                    st.error("Report generation failed.")
                    st.code(result.stderr or result.stdout)
                else:
                    html = html_output.read_text(encoding="utf-8")
                    excel_bytes = excel_output.read_bytes()

                    if save_to_archive:
                        archive_path = REPORT_DIR / f"{report_title}.html"
                        archive_path.write_text(html, encoding="utf-8")
                        st.success(f"Saved to history: {archive_path.name}")

                    st.download_button(
                        "Download Excel Report",
                        excel_bytes,
                        file_name=f"{report_title}.xlsx",
                    )

                    st.download_button(
                        "Download HTML Report",
                        html,
                        file_name=f"{report_title}.html",
                        mime="text/html",
                    )

                    components.html(html, height=1800, scrolling=True)
    else:
        st.info("Upload both Excel files to generate a dashboard.")


with archive_tab:
    st.subheader("Saved Dashboards")

    html_reports = sorted(REPORT_DIR.glob("*.html"), reverse=True)

    if not html_reports:
        st.info("No saved dashboards found.")
    else:
        selected_report = st.selectbox(
            "Select a dashboard",
            html_reports,
            format_func=lambda path: path.stem,
        )

        html = selected_report.read_text(encoding="utf-8")

        st.download_button(
            "Download HTML Report",
            html,
            file_name=selected_report.name,
            mime="text/html",
        )

        components.html(html, height=1800, scrolling=True)