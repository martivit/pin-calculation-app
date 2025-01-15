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
    # Add MSNA Indicators
    doc.add_heading('MSNA indicators/variables by dimension', level=2)
    msna_indicators = parameters["msna_indicators"]
    for category, indicators in msna_indicators.items():
        if isinstance(indicators, dict):  # Nested categories
            # Main bullet for the category with bold formatting
            category_paragraph = doc.add_paragraph(style='List Bullet')
            category_run = category_paragraph.add_run(f"{category.replace('_', ' ').capitalize()}:")
            category_run.bold = True
            for description, indicator in indicators.items():
                # Sub-bullets for each indicator
                doc.add_paragraph(f"  - {description}: {indicator}", style='List Bullet 2')
        else:
            # Main bullet for simple categories with bold formatting
            category_paragraph = doc.add_paragraph(style='List Bullet')
            category_run = category_paragraph.add_run(f"{category.replace('_', ' ').capitalize()}: {indicators}")
            category_run.bold = True


    # Add Severity Classification
    doc.add_heading('Severity Classification used for this calculation', level=2)
    severity_classification = parameters["severity_classification"]
    for level, details in severity_classification.items():
        # Color for severity level titles
        color_map = {
            "severity_level_3": RGBColor(255, 165, 0),  # Light orange
            "severity_level_4": RGBColor(255, 140, 0),  # Darker orange
            "severity_level_5": RGBColor(255, 69, 0),   # Red-orange
        }

        # Add severity level heading
        severity_paragraph = doc.add_paragraph(style='List Bullet')
        severity_run = severity_paragraph.add_run(f"{level.replace('_', ' ').capitalize()}: ")
        severity_run.bold = True
        if level in color_map:
            severity_run.font.color.rgb = color_map[level]

        # Add description
        description = details["description"]
        if "details" in details:
            description = description.replace("disrupted due to:", f"disrupted due to {details['details']}")
        severity_paragraph.add_run(description)

        # Add examples as a numbered list
        if "examples" in details:
            for i, example in enumerate(details["examples"], start=1):
                example_paragraph = doc.add_paragraph(f"{i}. {example}", style='List Number')
                # Underline specific keywords (for demonstration purposes)
                if "school" in example.lower():
                    example_run = example_paragraph.runs[0]
                    example_run.underline = True

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
