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
    doc = docx.Document()
    doc.add_heading('Parameters Used as Input for the PiN Calculation', level=1)

    # Add General Information
    doc.add_heading('General Information', level=2)
    general_info = parameters["general_info"]
    for key, value in general_info.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add MSNA Indicators
    doc.add_heading('MSNA indicators/variables by dimension', level=2)
    msna_indicators = parameters["msna_indicators"]
    for category, indicators in msna_indicators.items():
        if isinstance(indicators, dict):  # Nested categories
            doc.add_paragraph(f"{category.replace('_', ' ').capitalize()}:", style='List Bullet')
            for description, indicator in indicators.items():
                doc.add_paragraph(f"  - {description}: {indicator}", style='List Bullet')
        else:
            doc.add_paragraph(f"{category.replace('_', ' ').capitalize()}: {indicators}", style='List Bullet')

    # Add Severity Classification
    doc.add_heading('Severity Classification used for this calculation', level=2)
    severity_classification = parameters["severity_classification"]
    for level, details in severity_classification.items():
        # Start with the description
        description = details["description"]

        # Handle `details1` and `details2` explicitly
        if "details1" in details and "details2" in details:
            description = description.replace("disrupted due to.", 
                                              f"disrupted due to {details['details1']} and {details['details2']}.")

        # Integrate `details` for other cases
        elif "details" in details:
            description = description.replace("disrupted due to:", 
                                              f"disrupted due to {details['details']}")

        doc.add_paragraph(f"{level.replace('_', ' ').capitalize()}: {description}", style='List Bullet')

        # Add examples as bullet points
        if "examples" in details:
            for example in details["examples"]:
                doc.add_paragraph(f"  - {example}", style='List Bullet')

    # Add Admin Unit
    doc.add_heading('Administrative Unit', level=2)
    admin_unit = parameters["admin_unit"]
    for key, value in admin_unit.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Add School Cycles
    doc.add_heading('School Cycles', level=2)
    school_cycles = parameters.get("school_cycles", {})
    # Report 'age_ranges' (vector_cycle) as-is
    age_ranges = school_cycles.get("age_ranges", [])
    doc.add_paragraph(f"Age Ranges: {age_ranges}", style='List Bullet')

    # Handle 'notes' gracefully
    notes = school_cycles.get("notes", "Not specified")
    doc.add_paragraph(f"Notes: {notes}", style='List Bullet')

    # Save the Word document to a BytesIO object
    from io import BytesIO
    doc_output = BytesIO()
    doc.save(doc_output)
    doc_output.seek(0)

    return doc_output
