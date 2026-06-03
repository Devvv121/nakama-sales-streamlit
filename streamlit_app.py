import streamlit as st
import pandas as pd

st.set_page_config(page_title="Nakama Sales Dashboard", layout="wide")

st.title("Nakama Sales Dashboard")

sales_file = st.file_uploader("Upload Sales Orders Excel", type=["xlsx", "xls"])
products_file = st.file_uploader("Upload Products Excel", type=["xlsx", "xls"])

if sales_file and products_file:
    sales = pd.read_excel(sales_file)
    products = pd.read_excel(products_file)

    st.subheader("Sales Orders Preview")
    st.dataframe(sales.head(100), use_container_width=True)

    st.subheader("Products Preview")
    st.dataframe(products.head(100), use_container_width=True)

    st.success("Files loaded successfully.")
else:
    st.info("Upload both Excel files to start.")