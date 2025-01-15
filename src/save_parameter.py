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


def create_parameters_document(parameters, title="Parameters used as input for the PiN calculation"):
    """
    Creates a Word document listing the provided parameters and their values.

    :param parameters: A dictionary where keys are parameter names and values are their corresponding values.
    :param title: Title for the Word document.
    :return: A BytesIO object containing the Word document.
    """
    # Initialize the Word document
    doc = docx.Document()

    # Add the main title
    title_paragraph = doc.add_paragraph(title)
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_paragraph.add_run()
    title_run.font.size = Pt(16)
    title_run.bold = True
    title_run.font.name = "Calibri"

    # Add a blank line for spacing
    doc.add_paragraph()

    # Add a section for the parameters
    doc.add_heading("Parameters and Values", level=2)
    for param, value in parameters.items():
        # Add each parameter and its value
        param_paragraph = doc.add_paragraph()
        param_paragraph.add_run(f"{param}: ").bold = True
        param_paragraph.add_run(str(value))

    # Save the document to a BytesIO object
    doc_output = BytesIO()
    doc.save(doc_output)
    doc_output.seek(0)

    return doc_output