import streamlit as st
import pandas as pd
import io
import os
import base64


def to_excel(data):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output)
    data.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.save()
    processed_data = output.getvalue()

    return processed_data


def get_table_download_link(data):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(data)
    b64 = base64.b64encode(val)  # val looks like b'...'

    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="extract.xlsx">Download csv file</a>'  # decode b'abc' => abc


def parse_dataframe_windows7(data):

    data = data[
        (data["Objet du partage"].isin(["Caisse", "Comptabilité"]))
        | (data["Commentaire (Commentaire)"].str.startswith("DDR3", na=False))
        | (data["Commentaire (Commentaire)"].str.startswith("CAISSE", na=False))
        | (data["Commentaire (Commentaire)"].str.startswith("PANINI", na=False))
    ]

    return data


def write_save_result_excel(data, path_to_excel_result="classeur_results.xlsx"):

    df = parse_dataframe_windows7(data)

    # save resultst in new excel
    data.to_excel(
        path_to_excel_result, sheet_name="Sheet2", float_format="%.0f", index=False
    )
    with pd.ExcelWriter(path_to_excel_result, mode="a", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet3", float_format="%.0f", index=False)

    return df


def get_binary_file_downloader_html(bin_file, file_label="File"):
    with open(bin_file, "rb") as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'

    return href


if __name__ == "__main__":

    st.title("Boomscud first app to analyse Excel 😎")

    # Write DataFrame
    st.write("*Excel du parc informatique*")

    try:
        uploaded_file = st.file_uploader("Choose a file")
        data = pd.read_excel(uploaded_file, header=1)

        # Checkboxes to show/hide data
        if st.checkbox("Show data parc info"):
            st.write("Voila les 10 premieres lignes du parc total")
            st.write(data)

        if uploaded_file is not None:
            # read xls or xlsx
            df = write_save_result_excel(data)

            if st.checkbox("Show data windows7"):
                st.write("Voila les 10 premieres lignes du parc Windows 7")
                st.write(df)

                # st.markdown(get_binary_file_downloader_html("res.xlsx", "classeur excel Windows 7"), unsafe_allow_html=True)
            st.markdown(get_table_download_link(df), unsafe_allow_html=True)

        else:
            st.warning("you need to upload your Excel file (.xlsx)")

    except:
        st.warning("you need to upload your Excel file (.xlsx) / or computation failed")

    # # Show Progress
    # "Starting a long computation..."
    # # Add a placeholder
    # latest_iteration = st.empty()
    # bar = st.progress(0)

    # for i in range(100):
    #     # Update the progress bar with each iteration.
    #     latest_iteration.text(f"Iteration {i+1}")
    #     bar.progress(i + 1)
    #     time.sleep(0.01)

    # "...and now we're done!"

    # Or even better, call Streamlit functions inside a "with" block:
