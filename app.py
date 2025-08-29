# app.py

import streamlit as st
import pandas as pd
import os
from tbca_ppt_generator import generate_ppt

# Constants
DEFAULT_CSV_PATH = "data/mock_brandbook_monthly.csv"
OUTPUT_PPT_PATH = "output/TBCA_oral_care_pepsodent_core_2024.pptx"

st.set_page_config(page_title="PowerPoint Automation Demo", layout="wide")

st.title("📊 PowerPoint Deck Generator (Marketing Data)")
st.markdown("""
This app generates a branded PowerPoint presentation from monthly brand performance data.  
You can upload your own CSV (please ensure the template you use is similar to the demo dataset) or use the demo dataset provided.
""")

# --- Sidebar Options ---
st.sidebar.header("Upload Data")
uploaded_file = st.sidebar.file_uploader("""Upload CSV. \n
• The uploaded data will customize charts only up to **Slide 6** in this demo.  
• In real practice, you can customize **all charts** with your data. """, type="csv")

# Load data
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.success("✅ File uploaded successfully.")
else:
    df = pd.read_csv(DEFAULT_CSV_PATH)
    st.info("📎 Using default mock data.")

# Show raw data
st.subheader("📄 Data Preview")
st.dataframe(df)


# --- Generate PowerPoint ---
st.subheader("📤 Generate PowerPoint")
if st.button("Create Presentation"):
    # Save to default location for generator
    df.to_csv(DEFAULT_CSV_PATH, index=False)
    generate_ppt()
    st.success("✅ Presentation created!")

    # Show download button
    with open(OUTPUT_PPT_PATH, "rb") as file:
        st.download_button(
            label="📥 Download PowerPoint",
            data=file,
            file_name="TBCA_Brand_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
else:
    st.caption("Click to create and download your deck.")

# --- Footer ---
st.markdown("---")
st.markdown("Made with ❤️ for portfolio demonstration.")
