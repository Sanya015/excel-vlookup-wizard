import streamlit as st
from vlookup_processor import process_vlookup

st.set_page_config(page_title="Excel VLOOKUP App", layout="centered")

st.title("💡 Excel VLOOKUP Generator")
st.subheader("Upload 2 Excel files and auto-fill formulas!")

# Step 1: File Upload
file_a = st.file_uploader("📄 Upload Main Excel File", type=["xlsx"])
file_b = st.file_uploader("📄 Upload Lookup Excel File (O.xlsx)", type=["xlsx"])

# Step 2: User Inputs
sheet_choice = st.selectbox("Select sheet to process", ["UPPAbaby Order Form", "4Moms"])
last_row = st.number_input("Enter the last row to process", min_value=1, value=20)

# Pre-filled config
config = {
    "UPPAbaby Order Form": {"sheet_name": "UPPAbaby Order Form", "start_line": 11, "skip": [7,8,9]},
    "4Moms": {"sheet_name": "4Moms", "start_line": 8, "skip": None}
}

if st.button("⚙️ Process"):
    if file_a and file_b:
        st.info("Processing...")
        conf = config[sheet_choice]
        output_file = process_vlookup(file_a, file_b, conf["sheet_name"], conf["start_line"], last_row, conf["skip"])

        with open(output_file, "rb") as f:
            st.download_button("⬇️ Download Processed File", f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Please upload both files!")
