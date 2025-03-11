

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









    indicator_per_admin_status = {}
    # Assume category_data_frames is a dictionary of DataFrames, indexed by category
    for category, df in category_data_frames.items():
        # Ensure both DataFrames are ready to merge
        if category in pin_by_indicator_status_list:
            # Fetch the corresponding DataFrame from the grouped data
            grouped_df = pin_by_indicator_status_list[category]     
            # Merge on specified columns
            pop_group_df = pd.merge(grouped_df, df, on=[admin_var, pop_group_var])
            pop_group_df.columns = [str(col) for col in pop_group_df.columns]

            ##  Ensure `label_tot_population` is numeric before using it
            pop_group_df[label_tot_population] = pd.to_numeric(pop_group_df[label_tot_population], errors='coerce').fillna(0)

            ##  Rename columns
            pop_group_df = pop_group_df.rename(columns={
                pop_group_var: 'Population group',
                'sev3_indicator_teacher': label_perc_sev3_indicator_teacher,
                'sev3_indicator_hazard': label_perc_sev3_indicator_hazard,
                'sev3_indicator_access': label_perc_sev3_indicator_access,
                'sev4_indicator_idp': label_perc_sev4_indicator_idp,
                'sev5_indicator_occupation': label_perc_sev5_indicator_occupation,
                'sev4_aggravating_circumstances': label_perc_sev4_aggravating_circumstances,
                'sev5_aggravating_circumstances': label_perc_sev5_aggravating_circumstances
            })

            if 'Category' in pop_group_df.columns:
                del pop_group_df['Category']

            ##  Find matching columns for percentage & total number calculation
            total_columns = {}
            for col in pop_group_df.columns:
                if col not in ["Population group", admin_var]:  # Exclude these
                    if "severity level" in col and "# of children" not in col:
                        total_columns[col] = re.sub(r"% of children", "# of children", col)

            print(total_columns)            

            ##  Debug: Print column pairs to verify matching
            print("\n🔹 Matching Columns for Multiplication:")
            for perc_col, tot_col in total_columns.items():
                print(f"✔ {perc_col}  --->  {tot_col}")

            ##  Ensure percentage columns are numeric before multiplying
            for perc_col in total_columns.keys():
                pop_group_df[perc_col] = pd.to_numeric(pop_group_df[perc_col], errors='coerce').fillna(0)

            ##  Compute (ToT # children) values correctly
            for perc_col, tot_col in total_columns.items():
                if tot_col not in pop_group_df.columns:
                    pop_group_df[tot_col] = 0  # Ensure column exists

                # Extract values for debugging
                percentages = pop_group_df[perc_col]
                populations = pop_group_df[label_tot_population]
                computed_totals = (percentages * populations).round(0)

                #  Perform correct multiplication and rounding
                pop_group_df[tot_col] = computed_totals


            ##  Column Reordering
            all_columns = list(pop_group_df.columns)

            admin_cols = [admin_var, "Population group", "TotN"]
            # Extract severity levels and corresponding total columns dynamically
            severity_groups = {3: [], 4: [], 5: []}
            total_columns_map = {}

            for col in all_columns:
                if "severity level 3" in col and "# of children" not in col:
                    severity_groups[3].append(col)
                elif "severity level 4" in col and "# of children" not in col:
                    severity_groups[4].append(col)
                elif "severity level 5" in col and "# of children" not in col:
                    severity_groups[5].append(col)

                if "# of children" in col:
                    base_col = col.replace(" # of children", "")
                    total_columns_map[base_col] = col  # Map to its corresponding ToT column

            # Build ordered columns ensuring (ToT # children) comes immediately after its indicator
            final_columns = admin_cols
            for severity in [3, 4, 5]:  # Ordered severity levels
                for col in severity_groups[severity]:
                    final_columns.append(col)
                    if col in total_columns_map:  # Insert its total column immediately after
                        final_columns.append(total_columns_map[col])

            # Apply new order
            pop_group_df = pop_group_df[final_columns]

            #  Debugging: Print sample data to verify calculations
            print("\n📌 Sample Data After Calculation:")
            print(pop_group_df.columns)  # Show first few rows to verify correctness

            # Save modified DataFrame back into the dictionary under the category key
            indicator_per_admin_status[category] = pop_group_df
