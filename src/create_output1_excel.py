import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime  


int_2 = '2.0'
int_3 = '3.0'
int_4 = '4.0'
int_5 = '5.0'
label_perc2 = '% severity levels 1-2'
label_perc3 = '% severity level 3'
label_perc4 = '% severity level 4'
label_perc5 = '% severity level 5'
label_tot2 = '# severity levels 1-2'
label_tot3 = '# severity level 3'
label_tot4 = '# severity level 4'
label_tot5 = '# severity level 5'
label_perc_tot = '% Tot PiN (severity levels 3-5)'
label_tot = '# Tot PiN (severity levels 3-5)'
label_admin_severity = 'Area severity'
label_tot_population = 'TotN'

int_acc = 'access'
int_agg= 'aggravating circumstances'
int_lc = 'learning condition'
int_penv = 'protected environment'
int_out = 'Not in need'
label_perc_acc = '% Access'
label_perc_agg= '% Aggravating circumstances'
label_perc_lc = '% Learning conditions'
label_perc_penv = '% Protected environment'
label_perc_out = '% Not in need'
label_tot_acc = '# Access'
label_tot_agg= '# Aggravating circumstances'
label_tot_lc = '# Learning conditions'
label_tot_penv = '# Protected environment'
label_tot_out = '# Not in need'
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



indicator_patterns = [
    "severity (HPC2025) level 3 -- OoS children",
    "severity (HPC2025) level 3 -- in-school children",
    "severity (HPC2025) level 4 -- in-school children",
    "severity (HPC2025) level 4 -- OoS children",
    "severity (HPC2025) level 5 -- OoS children",
    "severity (HPC2025) level 5 -- in-school children",
]



def write_with_block_headers(df,  blocks):
    wb = Workbook()
    ws = wb.active
    max_col = ws.max_column

    for i in range(1, max_col +1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = 12
    for i in range(15, 19):  # 15=O, 16=P, … 21=U
            ws.column_dimensions[get_column_letter(i)].width = 15
    for i in range(19, 23):  # 15=O, 16=P, … 21=U
            ws.column_dimensions[get_column_letter(i)].width = 15        
        # 1) write a blank row 1 and then your header+data starting at row 2


    ws.append([])

    for raw_row in dataframe_to_rows(df, index=False, header=True):
        clean_row = []
        for cell in raw_row:
            # leave None, str, int, float, bool, date/datetime alone...
            if cell is None or isinstance(cell, (str, int, float, bool, datetime.date, datetime.datetime)):
                clean_row.append(cell)
            else:
                # everything else gets coerced to its string repr
                clean_row.append(str(cell))
        ws.append(clean_row)

    max_row = ws.max_row
    max_col = ws.max_column

    wrap_center = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )
    # only wrap & center rows 1 and 2, columns A→Z (or up to max_col)
    for r in (1, 2):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).alignment = wrap_center

    # 2) apply a very thin 'hair' border around every data cell
    hair = Side(style='thin', color='000000')
    hair_border = Border(left=hair, right=hair, top=hair, bottom=hair)
    for r in range(2, max_row+1):
        for c in range(1, max_col+1):
            ws.cell(row=r, column=c).border = hair_border

    # prepare styles for headers & block separators
    thin = Side(style='thin', color='000000')
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    thick = Side(style='thick', color='000000')

    # 3) loop your blocks
    for blk in blocks:
        cols, title, color = blk['columns'], blk['title'], blk['color']
        # skip defensive: no columns
        if not cols:
            continue
        # double-check they exist (should already be true after prune)
        idxs = [df.columns.get_loc(c) + 1 for c in cols if c in df.columns]
        if not idxs:
            continue
        start_col, end_col = min(idxs), max(idxs)

        # 3a) style & set the header‑cell before merging
        master = ws.cell(row=1, column=start_col)
        master.value = title
        master.fill = PatternFill(start_color=color,
                                  end_color=color,
                                  fill_type='solid')
        master.font = Font(bold=True)
        master.alignment = Alignment('center','center')
        master.border = header_border

        # 3b) apply header border & fill to all header cells in block
        for col in range(start_col, end_col+1):
            h = ws.cell(row=1, column=col)
            h.fill = PatternFill(start_color=color,
                                 end_color=color,
                                 fill_type='solid')
            h.alignment = Alignment('center','center')
            h.border = header_border

        # 3c) merge the header
        ws.merge_cells(start_row=1, start_column=start_col,
                       end_row=1,   end_column=end_col)

        # 3d) fill every cell in this block’s columns (data rows)
        fill = PatternFill(start_color=color,
                           end_color=color,
                           fill_type='solid')
        for col in range(start_col, end_col+1):
            for row in range(2, max_row+1):
                ws.cell(row=row, column=col).fill = fill

        special_col = 'Reference area Admin Pcode (for extrapolation)'
        if special_col in df.columns:
            special_idx = df.columns.get_loc(special_col) + 1
            special_fill = PatternFill(start_color='4af298',  # your new hex
                                    end_color='4af298',
                                    fill_type='solid')
            # header cell override:
            hdr_cell = ws.cell(row=1, column=special_idx)
            hdr_cell.fill = special_fill
            # data rows override:
            for r in range(2, max_row+1):
                ws.cell(row=r, column=special_idx).fill = special_fill        

        # 3e) draw a thick right‑hand border on this block’s edge
        for row in range(1, max_row+1):
            cell = ws.cell(row=row, column=end_col)
            b = cell.border
            cell.border = Border(
                left=b.left, top=b.top, bottom=b.bottom,
                right=thick
            )

    # 4) bump up header row height a bit
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 60

    return wb




def prune_blocks(df: pd.DataFrame, blocks):
    """Keep only columns that exist in df. Skip blocks with no remaining columns.
       Returns (pruned_blocks, missing_by_block)."""
    pruned = []
    missing = {}
    for blk in blocks:
        want = blk.get('columns', [])
        present = [c for c in want if c in df.columns]
        if present:
            pruned.append({**blk, 'columns': present})
        else:
            # whole block missing -> record for optional logging
            missing[blk.get('title', 'Untitled Block')] = [c for c in want if c not in df.columns]
    return pruned, missing



def create_output1_user(output1_platform):
    
   
    indicator_cols = [
        col for col in output1_platform.columns
        if any(pat in col for pat in indicator_patterns)
    ]

    
    blocks = [
        { 'columns': ['Admin', 'Admin Pcode'],
        'title': '2026 HPC scope',
        'color': colors['white'] },
        { 'columns': ['TotN',
            '% severity levels 1-2', '# severity levels 1-2',
            '% severity level 3',     '# severity level 3',
            '% severity level 4',     '# severity level 4',
            '% severity level 5',     '# severity level 5',
            '% Tot PiN (severity levels 3-5)',
            '# Tot PiN (severity levels 3-5)', 'Area severity'
        ],
        'title': 'PiN by severity, output platform with available 2025 data',
        'color': colors['light_beige'] },
        { 'columns': [
            "Education needs stable (tick 'X' if no change from last year)",
            "Education needs decreased (tick 'X' if improved from last year)",
            "Education needs worsened (tick 'X' if worsened from last year)"
        ],
        'title': 'Evolution (2024 --> 2025) of the education needs',
        'color': colors['light_yellow'] },
        { 'columns': [
            'Reference area Admin Pcode (for extrapolation)',
            'Suggested reference area for extrapolation',
            'Suggested reference area for extrapolation, LABEL',
            'Alternative reference areas',
            "Alternative reference areas, LABELS"
        ],
        'title': 'Identification of the proxy area',
        'color': colors['light_pink'] },
        { 'columns': [
            '% severity (HPC2025) levels 1-2',
            '% severity (HPC2025) level 3',
            '% severity (HPC2025) level 4',
            '% severity (HPC2025) level 5',
            '% Tot PiN (severity (HPC2025) levels 3-5)'
        ],
        'title': 'PiN by severity, HPC 2025',
        'color': colors['gray'] },
        {  "columns": indicator_cols,
        'title': 'PiN by indicator, HPC 2025',
        'color': colors['stratagray'] },
        { 'columns': [
            'Attacks on Schools','Attacks on Universities','Military Occupation of Education facility',
            'Arson attack on education facility','Forced Entry into education facility',
            'Damage/Destruction To Ed facility Event','Educators Killed','Educators Injured',
            'Educators Kidnapped','Educators Arrested','Students Attacked in School',
            'Students Killed','Students Injured','Students Kidnapped','Students Arrested',
            'Sexual Violence Affecting School Age Children','event_count','Date',
            'Event Description','Location of event','Reported Perpetrator',
            'Reported Perpetrator Name','Weapon Carried/Used','Type of education facility',
            'Known Educators Kidnap Or Arrest Outcome','Known Student Kidnap Or Arrest Outcome','ADM3_PCODE'
        ],
        'title': 'Insecurity insight',
        'color': colors['light_blue'] },
        { 'columns': [
            'admin1','admin2','fatalities','event_count_evt2','ADM3_PCODE_evt2',
            'event_type','sub_event_type','actor1','assoc_actor_1','actor2',
            'assoc_actor_2','notes','event_date'
        ],
        'title': 'ACLED',
        'color': colors['dark_orange'] },
        { 'columns': [
            'ratio IDP/ToTN - HPC2025','ratio IDP/ToTN - HPC2026'
        ],
        'title': 'IDP ratios',
        'color': colors['light_orange'] }
    ]

    blocks, missing = prune_blocks(output1_platform, blocks)
    if missing:
        print("⚠️ Missing columns by block:", missing)
    # 3) generate a styled Workbook
    wb = write_with_block_headers(output1_platform, blocks)

    # 4) save it into a BytesIO and return that
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output