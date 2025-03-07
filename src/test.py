

######################################################################### no ocha data

if no_ocha_data:
    (severity_admin_status_list, dimension_admin_status_list,
    indicator_per_admin_status,
    country_label) = calculatePIN_NO_OCHA_2025 (country, edu_data_severity, household_data, choice_data, survey_data,mismatch_ocha_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin,
                                                                                    selected_language= selected_language)

    indicator_output = create_indicator_output_no_ocha(country_label, indicator_per_admin_status, admin_var=admin_var)

    
    if selected_language == "English":
        doc_parameter_output = generate_word_document(parameters)

    if selected_language == "French":
        doc_parameter_output = generate_word_document_FR(parameters_FR)

    zip_file_name = f"PiN_by_indicator_Documents_{country_label}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    zip_file = create_zip_file_no_ocha(country_label, indicator_output,  doc_parameter_output)


    

    # Create a single download button for the ZIP file
    if st.download_button(
        label=translations["download_all"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip"
    ):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")

        try:
            repo_name = "martivit/pin-calculation-app"
            branch_name = "develop_2025"

            # File paths in the repository
            file_path_in_repo_excel = f"platform_PiN_output/{country}/PiN_by_indicator_results_{country}_{timestamp}.xlsx"
            file_path_in_repo_doc = f"platform_PiN_output/{country}/PiN_parameters_{country}_{timestamp}.docx"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            # Try uploading both files
            try:
                pr_url_excel = upload_to_github(
                    file_content=indicator_output.getvalue(),
                    file_name=file_path_in_repo_excel,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN results no ocha (Excel) for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Excel file to GitHub: {e}")

            try:
                pr_url_doc = upload_to_github(
                    file_content=doc_parameter_output.getvalue(),
                    file_name=file_path_in_repo_doc,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN parameters (Word) for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")

            # Display success messages only if files were successfully uploaded
            if pr_url_excel:
                st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")
            if pr_url_doc:
                st.success(f"Word document uploaded to GitHub successfully! [View File]({pr_url_doc})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass
 
    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])







    # Create an in-memory BytesIO buffer to hold the Excel file
    excel_pin = BytesIO()

    # Create an Excel writer object and write the DataFrames to it
    with pd.ExcelWriter(excel_pin, engine='xlsxwriter') as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in severity_admin_status_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    # Set the buffer position to the start
    excel_pin.seek(0)

    # Create a download button for the Excel file in Streamlit
    st.download_button(
        label="Download PiN percentages by admin and by population group",
        data=excel_pin,
        file_name=f"PiN_percentages_{country_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


