import pandas as pd
from fuzzywuzzy import process
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell
import docx
from docx.shared import Pt, RGBColor, Inches, Cm
import matplotlib.pyplot as plt
from io import BytesIO
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from matplotlib.text import Text
from docx.enum.table import WD_TABLE_ALIGNMENT
import matplotlib as mpl
from docx import Document
from datetime import datetime




def generate_parameters(st_session_state):
    """
    Generate the parameters dictionary for PiN calculation.
    Natural hazard & additional indicators are placed under
    learning conditions (sev=3) or protected environment (sev=4),
    and default to 'no_indicator' when not selected.
    """

    # --- read values from session ---
    country = st_session_state.get('country')

    # Natural hazard: column name + severity
    hazard_col = (
        st_session_state.get('selected_disruption_natural_hazard_column')
        #or st_session_state.get('natural_hazard_disruption_var')
        or 'no_indicator'
    )
    hazard_sev = st_session_state.get('natural_hazard_disruption_severity', None)

    # Additional indicator: last picked column + severity (only if toggle is ON)
    additional_col = st_session_state.get('additional_indicator_last_var') 
    additional_sev = st_session_state.get('additional_indicator_last_severity', None) 
    if not additional_col:
        additional_col = 'no_indicator'

    # Other core indicators
    access_col   = st_session_state.get('access_var')
    teacher_col  = st_session_state.get('teacher_disruption_var')
    idp_col      = st_session_state.get('idp_disruption_var')
    armed_col    = st_session_state.get('armed_disruption_var')
    barrier_col  = st_session_state.get('barrier_var')

    # Place hazard/additional into the right buckets (learning vs protected)
    hazard_in_learning   = hazard_col if (hazard_col != 'no_indicator' and hazard_sev == 3) else 'no_indicator'
    hazard_in_protected  = hazard_col if (hazard_col != 'no_indicator' and hazard_sev == 4) else 'no_indicator'

    additional_in_learning  = additional_col if (additional_col != 'no_indicator' and additional_sev == 3) else 'no_indicator'
    additional_in_protected = additional_col if (additional_col != 'no_indicator' and additional_sev == 4) else 'no_indicator'

    # Build the dicts for the two dimensions
    learning_block = {
        "Education disrupted due to teacher absences": teacher_col,
        "Education disrupted due to natural hazard": hazard_in_learning,
        "Additional indicator (severity 3)": additional_in_learning,
    }

    protected_block = {
        "Education disrupted due to school being used as IDP shelter": idp_col,
        "Education disrupted due to school being occupied by armed groups": armed_col,
        "Education disrupted due to natural hazard": hazard_in_protected,
        "Additional indicator (severity 4)": additional_in_protected,
    }

    # Build severity classification section
    severity3 = {
        "description": "OoS children who do NOT endure aggravating circumstances or in-school children whose education was disrupted due to:",
        "ind1 in-school": teacher_col,
        "ind2 in-school (hazard, sev3)": hazard_in_learning,
        "ind3 in-school (additional, sev3)": additional_in_learning,
    }

    severity4 = {
        "description": "In-school children whose education disrupted due to (ind in-school) or OoS facing the following aggravating circumstances.",
        "ind in-school (idp shelter)": idp_col,
        "ind in-school (hazard, sev4)": hazard_in_protected,
        "ind in-school (additional, sev4)": additional_in_protected,
        "aggravating circumstances": st_session_state.get('selected_severity_4_barriers', []),
    }

    severity5 = {
        "description": "In-school children whose education disrupted due to (ind in-school) or OoS facing the following aggravating circumstances.",
        "ind in-school": armed_col,
        "aggravating circumstances": st_session_state.get('selected_severity_5_barriers', []),
    }

    parameters = {
        "general_info": {
            "country": country,
            "date_calculation": datetime.now().strftime("%d/%m/%Y %H:%M")
        },
        "msna_indicators_per_PiN_dimension": {
            "access": access_col,
            "learning condition": learning_block,
            "protected environment": protected_block,
            "aggravating_circumstances": barrier_col,
        },
        "severity_classification": {
            "severity level 3": severity3,
            "severity level 4": severity4,
            "severity level 5": severity5,
        },
        "admin_unit": {
            "HNO unit of analysis": st_session_state.get('admin_var'),
            "mismatch admin": st_session_state.get('mismatch_admin', False),
        },
        "school_cycles": {
            "age_ranges": st_session_state.get('vector_cycle'),
        }
    }

    return parameters





def generate_word_document(parameters):
    # Initialize the Word document
    doc = docx.Document()
    doc.add_heading('Parameters Used as Input for the PiN Calculation', level=1)

    # Add General Information
    doc.add_heading('General Information', level=2)
    general_info = parameters["general_info"]
    for key, value in general_info.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add MSNA Indicators
    # Add MSNA Indicators
    doc.add_heading('MSNA indicators/variables by dimension', level=2)
    msna_indicators = parameters["msna_indicators_per_PiN_dimension"]
    for category, indicators in msna_indicators.items():
        if isinstance(indicators, dict):  # Nested categories
            # Main bullet for the category with bold formatting
            category_paragraph = doc.add_paragraph(style='List Bullet')
            category_run = category_paragraph.add_run(f"{category.replace('_', ' ').capitalize()}:")
            category_run.bold = True
            for description, indicator in indicators.items():
                # Sub-bullets for each indicator
                doc.add_paragraph(f"      {description}: {indicator}", style='List Bullet 2')
        else:
            # Main bullet for simple categories with bold formatting
            category_paragraph = doc.add_paragraph(style='List Bullet')
            category_run = category_paragraph.add_run(f"{category.replace('_', ' ').capitalize()}: {indicators}")
            category_run.bold = True


        # Add Severity Classification
        doc.add_heading('Severity Classification used for this calculation', level=2)
        severity_classification = parameters["severity_classification"]

        color_map = {
            "severity level 3": RGBColor(255, 165, 0),  # Light orange
            "severity level 4": RGBColor(255, 140, 0),  # Darker orange
            "severity level 5": RGBColor(255, 69, 0),   # Red-orange
        }

        def list_inds(details, prefix="ind"):
            """Return a list of indicator values, skipping 'no_indicator'."""
            return [v for k, v in details.items() if k.lower().startswith(prefix) and isinstance(v, str) and v != 'no_indicator']

        for level, details in severity_classification.items():
            severity_paragraph = doc.add_paragraph(style='List Bullet')
            severity_run = severity_paragraph.add_run(f"{level.replace('_', ' ').capitalize()}: ")
            severity_run.bold = True
            if level in color_map:
                severity_run.font.color.rgb = color_map[level]

            # Description
            description = details.get("description", "")
            if "In-school children" in description:
                # Clean and more readable structure
                severity_paragraph.add_run("In-school children whose education was disrupted due to ")

                # Add in-school indicators
                inds = list_inds(details, prefix="ind")
                if inds:
                    for i, val in enumerate(inds):
                        r = severity_paragraph.add_run(val)
                        r.bold = True
                        if i < len(inds) - 1:
                            severity_paragraph.add_run(" and ")
                    severity_paragraph.add_run(".")

                # Add aggravating circumstances (new formatting)
                if "aggravating circumstances" in details and details["aggravating circumstances"]:
                    doc.add_paragraph("or OoS facing the following aggravating circumstances:", style='List Bullet 2')
                    for example in details["aggravating circumstances"]:
                        example_paragraph = doc.add_paragraph(style='List Bullet 2')
                        example_paragraph.add_run(f"      {example}")

            else:
                # For severity 3 (same style as before)
                severity_paragraph.add_run(description + " ")
                inds = list_inds(details, prefix="ind")
                if inds:
                    for i, val in enumerate(inds):
                        r = severity_paragraph.add_run(val)
                        r.bold = True
                        if i < len(inds) - 1:
                            severity_paragraph.add_run(" and ")
                    severity_paragraph.add_run(".")

        # Add Admin Unit
        doc.add_heading('Administrative Unit', level=2)
        admin_unit = parameters["admin_unit"]
        for key, value in admin_unit.items():
            doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

        # Add School Cycles
        doc.add_heading('School Cycles', level=2)
        school_cycles = parameters.get("school_cycles", {})
        age_ranges = school_cycles.get("age_ranges", [])
        doc.add_paragraph(f"Age Ranges: {age_ranges}", style='List Bullet')

        # Save the Word document to a BytesIO object
        doc_output = BytesIO()
        doc.save(doc_output)
        doc_output.seek(0)

        return doc_output