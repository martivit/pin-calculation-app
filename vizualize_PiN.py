import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell

int_2 = '2.0'
int_3 = '3.0'
int_4 = '4.0'
int_5 = '5.0'
label_perc2 = '% 1-2'
label_perc3 = '% 3'
label_perc4 = '% 4'
label_perc5 = '% 5'
label_tot2 = '# 1-2'
label_tot3 = '# 3'
label_tot4 = '# 4'
label_tot5 = '# 5'
label_perc_tot = '% Tot PiN (3+)'
label_tot = '# Tot PiN (3+)'
label_admin_severity = 'Area severity'
label_tot_population = 'TotN'

int_acc = 'access'
int_agg= 'aggravating circumstances'
int_lc = 'learning condition'
int_penv = 'protected environment'
int_out = 'not falling within the PiN dimensions'
label_perc_acc = '% Access'
label_perc_agg= '% Aggravating circumstances'
label_perc_lc = '% Learning conditions'
label_perc_penv = '% Protected environment'
label_perc_out = '% Not falling within the PiN dimensions'
label_tot_acc = '# Access'
label_tot_agg= '# Aggravating circumstances'
label_tot_lc = '# Learning conditions'
label_tot_penv = '# Protected environment'
label_tot_out = '# Not falling within the PiN dimensions'

label_dimension_perc_tot = '% Tot in PiN Dimensions'
label_dimension_tot = '# Tot in PiN Dimensions'

label_dimension_tot_population = 'TotN'



# Define the colors
colors = {
    "light_beige": "FFF2CC",
    "light_orange": "F4B183",
    "dark_orange": "ED7D31",
    "darker_orange": "C65911",
    "light_blue": "DDEBF7",
    "light_pink": "b3b389",
    "light_yellow": "ffffc5",
    "white": "FFFFFF",
    "bluepin": "004bb4",
    'gray': 'e0e0e0',
    'stratagray': 'F0F0F0'
}
# Define the columns to color
color_mapping = {
    label_perc2: colors["light_beige"],
    label_tot2: colors["light_beige"],
    label_perc3: colors["light_orange"],
    label_tot3: colors["light_orange"],
    label_perc4: colors["dark_orange"],
    label_tot4: colors["dark_orange"],
    label_perc5: colors["darker_orange"],
    label_tot5: colors["darker_orange"],
    label_perc_tot: colors["light_blue"],
    label_admin_severity: colors["light_blue"],
    label_tot: colors["light_blue"]
}
# Define the colors
colors_dimension = {
    "light_beige": "ebecc7",
    "light_orange": "c7ebec",
    "dark_orange": "c7d9ec",
    "darker_orange": "c7ecdb",
    'darker2_orange':'b3d3d4',
    "light_blue": "DDEBF7",
    "light_pink": "b3b389",
    "light_yellow": "ffffc5",
    "white": "FFFFFF",
    "bluepin": "004bb4",
    'gray': 'e0e0e0',
    'stratagray': 'F0F0F0'
}
# Define the columns to color
color_mapping_dimension = {
    label_perc_out: colors_dimension["light_beige"],
    label_tot_out: colors_dimension["light_beige"],
    label_perc_acc: colors_dimension["light_orange"],
    label_tot_acc: colors_dimension["light_orange"],
    label_perc_agg: colors_dimension["dark_orange"],
    label_tot_agg: colors_dimension["dark_orange"],
    label_perc_lc: colors_dimension["darker_orange"],
    label_tot_lc: colors_dimension["darker_orange"],
    label_perc_penv: colors_dimension["darker2_orange"],
    label_tot_penv: colors_dimension["darker2_orange"],
    label_dimension_perc_tot: colors_dimension["light_blue"],
    label_dimension_tot: colors_dimension["light_blue"]
}

alignment_columns = list(color_mapping.keys())

def apply_formatting(workbook, color_mapping, alignment_columns, colors, admin_var):
    for ws in workbook.worksheets:
        # Add empty rows at the top
        ws.insert_rows(1, 4)
        # Add empty columns to the left
        ws.insert_cols(1, 4)

        # Add the sheet name as a title in the first row
        title = ws.title
        if ws.title != "PiN TOTAL":
            title += " (5-17 y.o.)"
        max_col = ws.max_column
        ws.merge_cells(start_row=1, start_column=5, end_row=1, end_column=max_col)
        title_cell = ws.cell(row=1, column=5)
        title_cell.value = title
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')

        # Set column widths based on content
        for column in ws.columns:
            max_length = 0
            column_letter = None
            for cell in column:
                if cell.value and not isinstance(cell, MergedCell):
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                        column_letter = cell.column_letter  # Get the column letter
                    except Exception as e:
                        print(f"Error: {e}, Cell: {cell}")

            if column_letter:
                adjusted_width = max_length + 2
                ws.column_dimensions[column_letter].width = adjusted_width

        # Bold specific columns
        for row in ws.iter_rows(min_row=6, max_col=ws.max_column, max_row=ws.max_row):  # Start from the data row
            for cell in row:
                col_name = ws.cell(row=5, column=cell.column).value  # Row 5 contains the headers
                if col_name in [label_perc_tot, label_admin_severity, label_tot, admin_var]:
                    cell.font = Font(bold=True)  # Apply bold font

        # Check if the worksheet is "PiN TOTAL"
        if ws.title == "PiN TOTAL":
            for row in ws.iter_rows(min_row=5, max_col=ws.max_column, max_row=ws.max_row):
                for cell in row:
                    col_name = ws.cell(row=5, column=cell.column).value
                    if col_name in alignment_columns:
                        cell.alignment = Alignment(horizontal='right', vertical='center')

            # Iterate through the rows starting from the first data row
            for row in ws.iter_rows(min_row=5, max_col=ws.max_column, max_row=ws.max_row):
                strata_value = row[4].value  # 'Strata' column should be the 5th column after inserting 4 empty columns

                # Determine the fill color based on 'Strata' value
                if strata_value == "TOTAL (5-17 y.o.)":
                    fill_color = colors["bluepin"]
                    for cell in row:
                        cell.font = Font(color=colors["white"], bold=True)  # Set text color to white and bold
                elif strata_value in ["Girls", "Boys"]:
                    fill_color = colors["gray"]
                elif strata_value in ['Female (MSNA)', 'Male (MSNA)']:
                    fill_color = colors["stratagray"]
                elif strata_value == "ECE (5 y.o.)":
                    fill_color = colors["light_pink"]
                elif "school" in strata_value.lower():
                    fill_color = colors["light_yellow"]
                elif strata_value == "Strata":
                    fill_color = colors["white"]
                elif "disability" in strata_value.lower():
                    fill_color = colors["white"]
                else:
                    fill_color = colors["light_blue"]

                # Apply fill color to the entire row if a color is determined
                if fill_color:
                    for cell in row:
                        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

                # Set first four columns to white
                for cell in row[:4]:
                    cell.fill = PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid")
                    cell.border = None  # No border for the first four columns

        else:
            # Apply formatting to the headers
            for cell in ws[5]:  # Header row is now the 5th row
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Apply color to specific columns and make borders visible
            for row in ws.iter_rows(min_row=5, max_col=ws.max_column, max_row=ws.max_row):
                for cell in row:
                    col_name = ws.cell(row=5, column=cell.column).value
                    # Apply color based on mapping
                    if col_name in color_mapping:
                        cell.fill = PatternFill(start_color=color_mapping[col_name], end_color=color_mapping[col_name], fill_type="solid", patternType='solid')
                    # Apply alignment
                    if col_name in alignment_columns:
                        cell.alignment = Alignment(horizontal='right', vertical='center')

                    # Apply border, but skip the first four columns
                    if cell.column > 4:
                        if row[0].row == 5:  # Bold top border for header
                            cell.border = Border(
                                top=Side(style="medium"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )
                        elif cell == row[0]:  # Bold left border for each row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="medium"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )
                        elif cell == row[-1]:  # Bold right border for each row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="medium"),
                                bottom=Side(style="thin"),
                            )
                        elif row[0].row == ws.max_row:  # Bold bottom border for last row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="medium"),
                            )
                        else:
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )
                    else:
                        cell.fill = PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid")  # White background for the first four columns
                        cell.border = None  # No border for the first four columns

        # Apply borders around the table for "PiN TOTAL"
        if ws.title == "PiN TOTAL":
            for row in ws.iter_rows(min_row=5, max_col=ws.max_column, max_row=ws.max_row):
                for cell in row:
                    if cell.column > 4:
                        if row[0].row == 5:  # Bold top border for header
                            cell.border = Border(
                                top=Side(style="medium"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )
                        elif cell == row[0]:  # Bold left border for each row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="medium"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )
                        elif cell == row[-1]:  # Bold right border for each row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="medium"),
                                bottom=Side(style="thin"),
                            )
                        elif row[0].row == ws.max_row:  # Bold bottom border for last row
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="medium"),
                            )
                        else:
                            cell.border = Border(
                                top=Side(style="thin"),
                                left=Side(style="thin"),
                                right=Side(style="thin"),
                                bottom=Side(style="thin"),
                            )

    return workbook






#engine='openpyxl'
# Function to create an Excel file and return it as a BytesIO object


def create_output(dataframes, overview_df, overview_sheet_name, admin_var, dimension=True, ocha= True):
    output = BytesIO()
    with pd.ExcelWriter(output) as writer:
        # Only write the overview sheet if dimension is False
        if  ocha:
            overview_df.to_excel(writer, sheet_name=overview_sheet_name, index=False)
        
        # Write the category sheets
        for category, df in dataframes.items():
            sheet_name = f"{overview_sheet_name.split()[0]} -- {category}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    workbook = load_workbook(output)

    # Determine which color mapping and colors to use
    if dimension:
        workbook = apply_formatting(workbook, color_mapping_dimension, alignment_columns, colors_dimension, admin_var)
    else:
        workbook = apply_formatting(workbook, color_mapping, alignment_columns, colors, admin_var)
    
    formatted_output = BytesIO()
    workbook.save(formatted_output)
    formatted_output.seek(0)

    return formatted_output





