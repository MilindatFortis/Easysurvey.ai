import streamlit as st
from qre_formatter import process_qre
import os

# =====================================
# PAGE CONFIG
# =====================================

st.set_page_config(
    page_title="Survey Master AI Agent",
    layout="wide"
)

# =====================================
# TITLE
# =====================================

st.title("Survey Master AI Agent")

st.write(
    "Upload your QRE template and generate formatted output automatically."
)

# =====================================
# FILE UPLOADER
# =====================================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# =====================================
# PROCESS FILE
# =====================================

if uploaded_file is not None:

    st.success("File uploaded successfully!")

    # Create uploads folder
    os.makedirs("uploads", exist_ok=True)

    input_path = os.path.join(
        "uploads",
        uploaded_file.name
    )

    # Save uploaded file
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Run button
    if st.button("Run QRE Formatter"):

        try:

            with st.spinner("Processing file..."):

                output_file = process_qre(input_path)

            st.success("Formatting completed successfully!")

            # =====================================
            # DOWNLOAD BUTTON
            # =====================================

            with open(output_file, "rb") as f:

                st.download_button(
                    label="Download Formatted File",
                    data=f,
                    file_name=os.path.basename(output_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:

            st.error(f"Error: {str(e)}")