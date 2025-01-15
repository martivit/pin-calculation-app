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


def generate_word_document(parameters):
    # Initialize the Word document
    doc = Document()
    doc.add_heading('Parameters Used as Input for the PiN Calculation', level=1)

    # Add General Information
    doc.add_heading('General Information', level=2)
    general_info = parameters["general_info"]
    for key, value in general_info.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add MSNA Indicators
    doc.add_heading('MSNA Indicators', level=2)
    msna_indicators = parameters["msna_indicators"]
    for key, value in msna_indicators.items():
        if isinstance(value, dict):  # Handle nested dictionaries
            doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}:", style='List Bullet')
            for sub_key, sub_value in value.items():
                doc.add_paragraph(f"  - {sub_key.replace('_', ' ').capitalize()}: {sub_value}", style='List Bullet')
        else:
            doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add Severity Classification
    doc.add_heading('Severity Classification', level=2)
    severity_classification = parameters["severity_classification"]
    for level, details in severity_classification.items():
        doc.add_paragraph(f"{level.replace('_', ' ').capitalize()}:", style='List Bullet')
        doc.add_paragraph(f"  - Description: {details['description']}", style='List Bullet')
        if "examples" in details:
            doc.add_paragraph("  - Examples:", style='List Bullet')
            for example in details["examples"]:
                doc.add_paragraph(f"    * {example}", style='List Bullet')
        if "details" in details:
            doc.add_paragraph(f"  - Details: {details['details']}", style='List Bullet')

    # Add Admin Unit
    doc.add_heading('Administrative Unit', level=2)
    admin_unit = parameters["admin_unit"]
    for key, value in admin_unit.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add School Cycles
    doc.add_heading('School Cycles', level=2)
    school_cycles = parameters["school_cycles"]
    doc.add_paragraph(f"Age Ranges: {', '.join(map(str, school_cycles['age_ranges']))}", style='List Bullet')
    doc.add_paragraph(f"Notes: {school_cycles['notes']}", style='List Bullet')

    # Add Additional Information
    doc.add_heading('Additional Information', level=2)
    additional_info = parameters["additional_info"]
    for key, value in additional_info.items():
        if isinstance(value, list):  # Handle lists for selected_severity_4 and 5 barriers
            doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}:", style='List Bullet')
            for item in value:
                doc.add_paragraph(f"  - {item}", style='List Bullet')
        else:
            doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Save the Word document to a BytesIO object
    from io import BytesIO
    doc_output = BytesIO()
    doc.save(doc_output)
    doc_output.seek(0)

    return doc_output