import streamlit as st
import pandas as pd
import io
import os
import base64
import itertools


def to_excel(data, df_win7):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    data.to_excel(writer, sheet_name="Sheet1", index=False)
    df_win7.to_excel(writer, sheet_name="Sheet2", index=False)

    writer.save()
    processed_data = output.getvalue()

    return processed_data


def get_table_download_link(data, df_win7):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(data, df_win7)
    b64 = base64.b64encode(val)  # val looks like b'...'

    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="extract.xlsx">Download csv file</a>'  # decode b'abc' => abc


def apply_filter_windows7(data, filters):

    # build flat filter dictionnary
    filters_spread = []
    for col in filters.keys():
        filters_spread.append(list(itertools.product([col], filters[col])))

    flat_filters = [item for sublist in filters_spread for item in sublist]

    # build pandas msk (column startswith ...)
    msk = []
    for sublist in flat_filters:
        col, val = sublist
        msk.append((data[col].str.startswith(val, na=False)))

    # result
    final_msk = filter_many_or(msk)
    df = data[final_msk]

    return df


def write_save_result_excel(data, path_to_excel_result="classeur_results.xlsx"):

    df = parse_dataframe_windows7(data)

    return df


def filter_many_or(list_of_masks):
    aggregate_mask = list_of_masks[0]

    for mask in list_of_masks[1:]:
        aggregate_mask = aggregate_mask | mask

    return aggregate_mask


if __name__ == "__main__":

    st.title("Analyse excel du parc informatique 😎")

    st.subheader("Drop Excel file")

    try:
        uploaded_file = st.file_uploader("Choose a file")
        data = pd.read_excel(uploaded_file, header=1)

        # Checkboxes to show/hide data
        if st.checkbox("Show data parc info"):
            st.write("Voila les 10 premieres lignes du parc total")
            st.write(data)

        st.write("\n")

        if uploaded_file is not None:
            # read xls or xlsx

            # Columns to CHOOSE to apply filter
            cols_filtered = st.multiselect(
                "Colonnes possibles dans le fichier Excel", list(data.columns), key=0
            )
            st.write("Colonnes à filtrer:", cols_filtered)

            # value to choose in columns
            filters = {}
            defaults = {
                "Commentaire (Commentaire)": ["PANINI", "DDR3", "CAISSE"],
                "Objet du partage": ["Caisse", "Comptabilité"],
            }

            for col in cols_filtered:

                option = st.multiselect(
                    f"Valeurs possibles à filtrer dans la colonne : '{col}'",
                    list(data[col].unique()),
                    defaults[col],
                )
                filters[col] = option

                # st.write("You selected:", filters)

            df = apply_filter_windows7(data, filters)

            st.subheader("Result Excel")
            st.write("\n")

            if st.checkbox("Show data windows7"):
                st.write("Voila les 10 premieres lignes du parc Windows 7")
                st.write(df)

            # html link download excel
            st.markdown(get_table_download_link(data, df), unsafe_allow_html=True)

        else:
            st.warning("you need to upload your Excel file (.xlsx)")

    except:
        st.warning("Computation failed")

    # a = st.text_input("Valeurs à filtrer pour la colonne '{col}'")
    # st.write(a)
