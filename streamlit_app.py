import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Nakama Sales Dashboard", layout="wide")

st.title("Nakama Sales Dashboard")

sales_file = st.file_uploader("Upload Sales Orders Excel", type=["xlsx", "xls"])
products_file = st.file_uploader("Upload Products Excel", type=["xlsx", "xls"])

generator = Path(__file__).parent / "Generate-SalesBrandReport.py"

if sales_file and products_file:
    report_title = st.text_input("Report Title", value="Nakama Sales Report")

    if st.button("Generate Dashboard", type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)

            sales_path = tmp_path / sales_file.name
            products_path = tmp_path / products_file.name
            html_output = tmp_path / "report.html"
            excel_output = tmp_path / "report.xlsx"

            sales_path.write_bytes(sales_file.getbuffer())
            products_path.write_bytes(products_file.getbuffer())

            result = subprocess.run(
                [
                    sys.executable,
                    str(generator),
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

                with open(excel_output, "rb") as file:
                    st.download_button(
                        "Download Excel Report",
                        file,
                        file_name=f"{report_title}.xlsx",
                    )

                components.html(html, height=1800, scrolling=True)

else:
    st.info("Upload both Excel files to start.")