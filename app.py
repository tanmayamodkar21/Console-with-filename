import streamlit as st
import pandas as pd
import os

def load_data(file):
    # Read the file into a DataFrame
    df = pd.read_excel(file)
    
    # Extract the first three letters of the file name (excluding extension)
    name = os.path.splitext(file.name)[0][:3]
    
    # Replace specific initials with desired values
    if name.lower() == 'lal':
        office_name = 'LAX'
    elif name.lower() == 'new':
        office_name = 'NYC'
    else:
        office_name = name.upper()
    
    # Update the 'Office' column with the mapped initials or office codes
    df['Office'] = office_name.upper()
    return df

st.title('Excel Files Consolidation')

# Upload multiple files
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True, type=["xlsx", "xls"])

if uploaded_files:
    consolidated_df = pd.DataFrame()

    for file in uploaded_files:
        df = load_data(file)
        consolidated_df = pd.concat([consolidated_df, df], ignore_index=True)

    st.write(consolidated_df)

    # Option to download the consolidated file
    consolidated_df.to_excel("consolidated_file.xlsx", index=False)
    with open("consolidated_file.xlsx", "rb") as file:
        btn = st.download_button(
            label="Download consolidated file",
            data=file,
            file_name="consolidated_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
