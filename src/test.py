

        severity_admin_status_list = run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='severity_category',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        dimension_admin_status_list = run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        dimension_admin_status_in_need_list = run_mismatch_admin_analysis(in_need_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)      
        severity_female_list = run_mismatch_admin_analysis(female_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='severity_category',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  
        severity_male_list = run_mismatch_admin_analysis(male_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='severity_category',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  
        dimension_female_list = run_mismatch_admin_analysis(female_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)        
        dimension_male_list = run_mismatch_admin_analysis(male_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_ece_list = run_mismatch_admin_analysis(ece_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_primary_list = run_mismatch_admin_analysis(primary_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_secondary_list = run_mismatch_admin_analysis(secondary_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  



        indicator_access_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.access',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_teacher_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.teacher',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_hazard_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.hazard',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_idp_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.idp',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_occupation_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.occupation',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_barrier4_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.barrier4',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_barrier5_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='indicator.barrier5',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        indicator_barrier_list =  run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable=barrier_var,
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)





 #------    CORRECT PIN    -------            
        severity_admin_status = calculate_prop (df=df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'severity_category')
        #-------    CORRECT TARGETTING    -------          
        dimension_admin_status = calculate_prop (df=df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        ## subset in need
        dimension_admin_status_in_need = calculate_prop (df=in_need_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        # -------- GENDER DISAGGREGATION  ---------    
        severity_female = calculate_prop (df=female_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'severity_category')
        severity_male = calculate_prop (df=male_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'severity_category')
        dimension_female = calculate_prop (df=female_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        dimension_male = calculate_prop (df=male_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        # -------- SCHOOL-CYCLE DISAGGREGATION  ---------    
        dimension_ece = calculate_prop (df=ece_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        dimension_primary = calculate_prop (df=primary_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        dimension_secondary = calculate_prop (df=secondary_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')
        if not single_cycle:
            dimension_intermediate =  calculate_prop (df=intermediate_df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= 'dimension_pin')




            ## reducing the multiindex of the panda dataframe
        severity_admin_status_list = reduce_index(severity_admin_status, 0, pop_group_var)
        dimension_admin_status_list = reduce_index(dimension_admin_status, 0, pop_group_var)
        dimension_admin_status_in_need_list = reduce_index(dimension_admin_status_in_need,  0, pop_group_var) ## only who is in need we check the distriburion of need
        severity_female_list = reduce_index(severity_female, 0, pop_group_var)
        severity_male_list = reduce_index(severity_male, 0, pop_group_var)
        dimension_female_list = reduce_index(dimension_female, 0, pop_group_var)
        dimension_male_list = reduce_index(dimension_male, 0, pop_group_var)
        dimension_ece_list = reduce_index(dimension_ece, 0, pop_group_var)
        dimension_primary_list = reduce_index(dimension_primary, 0, pop_group_var)
        dimension_secondary_list = reduce_index(dimension_secondary, 0, pop_group_var)
        if not single_cycle: dimension_intermediate_list = reduce_index(dimension_intermediate, 0, pop_group_var)




    ## checking number of columns
    severity_needed_columns = [2.0, 3.0, 4.0, 5.0]
    dimension_needed_columns = ['access','aggravating circumstances', 'learning condition', 'protected environment']
    severity_admin_status_list = ensure_columns(severity_admin_status_list, severity_needed_columns)
    severity_female_list = ensure_columns(severity_female_list, severity_needed_columns)
    severity_male_list = ensure_columns(severity_male_list, severity_needed_columns)
    dimension_admin_status_list = ensure_columns(dimension_admin_status_list, dimension_needed_columns)
    dimension_admin_status_in_need_list = ensure_columns(dimension_admin_status_in_need_list, dimension_needed_columns)
    dimension_female_list = ensure_columns(dimension_female_list, dimension_needed_columns)
    dimension_male_list = ensure_columns(dimension_male_list, dimension_needed_columns)
    dimension_ece_list = ensure_columns(dimension_ece_list, dimension_needed_columns)
    dimension_primary_list = ensure_columns(dimension_primary_list, dimension_needed_columns)
    dimension_secondary_list = ensure_columns(dimension_secondary_list, dimension_needed_columns)
    if not single_cycle:    dimension_intermediate_list = ensure_columns(dimension_intermediate_list, dimension_needed_columns)