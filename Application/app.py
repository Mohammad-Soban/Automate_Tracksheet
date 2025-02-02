import streamlit as st
import os
from client_invoice import client_Invoice
from mt_incentives import MT_incentives
from Cumulative import Cumulative

def main():
    st.title("Client Invoice Generator")
    st.write("Upload an Excel file, and generate a processed Client Invoice file.")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])

    if uploaded_file:
        temp_file_path = os.path.join("temp", uploaded_file.name)
        os.makedirs("temp", exist_ok=True)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getvalue())

        st.success(f"File '{uploaded_file.name}' uploaded successfully!")

        if st.button("Generate Client Invoice"):
            st.write("Generating client invoice...")
            output_file, messages = client_Invoice(temp_file_path)

            if output_file:
                st.success("Client invoice file generated successfully!")
                if messages:
                    for msg in messages:
                        st.warning(msg)
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="Download Client Invoice",
                        data=f,
                        file_name="Client_Invoice.xls",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("Failed to generate the client invoice file.")
                for msg in messages:
                    st.error(msg)

        if st.button("Generate MT Incentives"):
            st.write("Generating MT incentives...")
            output_files, messages = MT_incentives(temp_file_path)

            if output_files:
                st.success("MT incentives files generated successfully!")
                if messages:
                    for msg in messages:
                        st.warning(msg)
                with open(output_files[0], "rb") as f1:
                    st.download_button(
                        label="Download Middle_1 File",
                        data=f1,
                        file_name="Middle_1.xls",
                        mime="application/vnd.ms-excel"
                    )
                with open(output_files[1], "rb") as f2:
                    st.download_button(
                        label="Download MT_Incentives File",
                        data=f2,
                        file_name="MT_Incentives.xls",
                        mime="application/vnd.ms-excel"
                    )
            else:
                st.error("Failed to generate the MT incentives files.")
                for msg in messages:
                    st.error(msg)

        if st.button("Generate Middle_2 and Cumulative"):
            st.write("Generating Middle_2 and Cumulative files...")
            cumulative_file, messages = Cumulative(temp_file_path)

            if cumulative_file:
                st.success("Middle_2 and Cumulative files generated successfully!")
                if messages:
                    for msg in messages:
                        st.warning(msg)
                with open(cumulative_file, "rb") as f:
                    st.download_button(
                        label="Download Cumulative File",
                        data=f,
                        file_name="Cumulative.xls",
                        mime="application/vnd.ms-excel"
                    )
            else:
                st.error("Failed to generate the Middle_2 and Cumulative files.")
                for msg in messages:
                    st.error(msg)

if __name__ == "__main__":
    main()