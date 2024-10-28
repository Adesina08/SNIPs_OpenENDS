import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO


# Function to process the upload and generate the matching Excel workbook
def process_open_end_collation(questionnaire_df, data_df, original_filename):
    # Filter "text" type columns from the first upload (questionnaire)
    text_filtered = questionnaire_df[questionnaire_df["type"] == "text"]
    name_columns = text_filtered["name"].tolist()

    # Ensure that the second upload's first row matches with 'name' columns
    data_columns = data_df.columns.tolist()
    matching_columns = [col for col in data_columns if col in name_columns]

    # Create a new workbook with the required columns and matched columns
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Get the required columns (instanceID, enu_id, state) and matching columns
        required_columns = ["instanceID", "enu_id", "state"]
        available_required = [col for col in required_columns if col in data_df.columns]

        # Combine required columns with matching columns
        columns_to_include = available_required + matching_columns
        final_df = data_df[columns_to_include]

        # Write to Excel
        final_df.to_excel(writer, index=False, sheet_name="Matched Data")

    # Generate output filename
    output_filename = original_filename.rsplit(".", 1)[0] + "_open_end_collation.xlsx"

    # Set up download link
    st.success("File processed successfully!")
    st.download_button(
        label="Download Collation Results",
        data=output.getvalue(),
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# Streamlit app
st.set_page_config(page_title="Survey Processing App", layout="wide")

# Sidebar for page navigation
page = st.sidebar.selectbox("Select a Page", ["Open End Collation", "Analysis"])

if page == "Open End Collation":
    st.title("Open End Collation")

    # Step 1: Upload the survey questionnaire (XLS or XLSX)
    st.subheader("Upload the Survey Questionnaire (XLS or XLSX format)")
    questionnaire_file = st.file_uploader(
        "Choose a questionnaire file", type=["xls", "xlsx"]
    )

    if questionnaire_file is not None:
        try:
            # Check file extension and read the file accordingly
            if questionnaire_file.name.endswith(".xls"):
                questionnaire_df = pd.read_excel(questionnaire_file, engine="xlrd")
            elif questionnaire_file.name.endswith(".xlsx"):
                questionnaire_df = pd.read_excel(questionnaire_file, engine="openpyxl")
            else:
                st.error("Unsupported file type. Please upload a .xls or .xlsx file.")
                questionnaire_df = None

        except Exception as e:
            st.error(f"Error reading questionnaire file: {e}")
            questionnaire_df = None

        if questionnaire_df is not None:
            if (
                "type" in questionnaire_df.columns
                and "name" in questionnaire_df.columns
            ):
                st.write("Survey Questionnaire uploaded successfully.")
                st.write(
                    "Focusing on 'text' and 'integer' columns in the 'type' column"
                )

                # Filter only "text" and "integer" columns from the 'type' column
                type_filtered_df = questionnaire_df[
                    questionnaire_df["type"].isin(["text", "integer"])
                ]
                st.write("Filtered Questionnaire Columns:")
                st.write(type_filtered_df)

                # Step 2: Upload the actual data (XLSX format)
                st.subheader("Upload the Survey Data (XLSX format)")
                data_file = st.file_uploader("Choose a data file", type=["xlsx"])

                if data_file is not None:
                    try:
                        data_df = pd.read_excel(data_file, engine="openpyxl")
                        st.write("Survey Data uploaded successfully.")

                        # Check for required columns
                        required_columns = ["instanceID", "enu_id", "state"]
                        missing_columns = [
                            col
                            for col in required_columns
                            if col not in data_df.columns
                        ]

                        if missing_columns:
                            st.error(
                                f"The uploaded data is missing the following required columns: {', '.join(missing_columns)}"
                            )
                        else:
                            # Step 3: Process the uploads and match data
                            process_open_end_collation(
                                type_filtered_df, data_df, data_file.name
                            )

                    except Exception as e:
                        st.error(f"Error reading survey data file: {e}")

elif page == "Analysis":
    st.title("Analysis Page")
    st.write(
        "This page is under construction. You can build your analysis features here."
    )
