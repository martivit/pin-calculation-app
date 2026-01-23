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
def generate_parameters_FR(st_session_state):
    """
    Génère le dictionnaire des paramètres (FR) pour le calcul PiN.
    L'aléa naturel et l'indicateur supplémentaire sont placés :
      - en "conditions d'apprentissage" si sévérité = 3
      - en "environnement protégé"   si sévérité = 4
    et tombent à 'no_indicator' s'ils ne sont pas sélectionnés.
    """

    # --- lecture des valeurs de session ---
    pays = st_session_state.get('country')

    # Aléa naturel : nom de colonne + sévérité
    hazard_col = (
        st_session_state.get('selected_disruption_natural_hazard_column')
        or 'no_indicator'
    )
    hazard_sev = st_session_state.get('natural_hazard_disruption_severity', None)

    # Indicateur supplémentaire : dernier choisi + sévérité (uniquement si activé)
    additional_enabled = st_session_state.get("additional_indicator_enable", False)
    additional_col = st_session_state.get('additional_indicator_last_var') if additional_enabled else None
    additional_sev = st_session_state.get('additional_indicator_last_severity') if additional_enabled else None
    if not additional_col:
        additional_col = 'no_indicator'
    # Indicateur supplémentaire : dernier choisi + sévérité (uniquement si activé)
    additional_2_enabled = st_session_state.get("additional_2_indicator_enable", False)
    additional_2_col = st_session_state.get('additional_2_indicator_last_var') if additional_2_enabled else None
    additional_2_sev = st_session_state.get('additional_2_indicator_last_severity') if additional_2_enabled else None
    if not additional_2_col:
        additional_2_col = 'no_indicator'    

    # Autres indicateurs clés
    acces_col   = st_session_state.get('access_var')
    prof_col    = st_session_state.get('teacher_disruption_var')
    idp_col     = st_session_state.get('idp_disruption_var')
    arme_col    = st_session_state.get('armed_disruption_var')
    barrier_col = st_session_state.get('barrier_var')

    # Répartition par sévérité
    hazard_en_apprent   = hazard_col if (hazard_col != 'no_indicator' and hazard_sev == 3) else 'no_indicator'
    hazard_env_protege  = hazard_col if (hazard_col != 'no_indicator' and hazard_sev == 4) else 'no_indicator'

    add_en_apprent      = additional_col if (additional_col != 'no_indicator' and additional_sev == 3) else 'no_indicator'
    add_env_protege     = additional_col if (additional_col != 'no_indicator' and additional_sev == 4) else 'no_indicator'
    add_2_en_apprent      = additional_2_col if (additional_2_col != 'no_indicator' and additional_2_sev == 3) else 'no_indicator'
    add_2_env_protege     = additional_2_col if (additional_2_col != 'no_indicator' and additional_2_sev == 4) else 'no_indicator'
    # Blocs par dimension
    bloc_apprentissage = {
        "Éducation perturbée en raison de l'absence des enseignants": prof_col,
        "Éducation perturbée en raison d'un aléa naturel": hazard_en_apprent,
        "Indicateur supplémentaire (sévérité 3)": add_en_apprent,
        "Indicateur supplémentaire2 (sévérité 3)": add_2_en_apprent,

    }

    bloc_protege = {
        "Éducation perturbée en raison de l'utilisation de l'école comme abri pour les PDI": idp_col,
        "Éducation perturbée en raison de l'occupation de l'école par des groupes armés": arme_col,
        "Éducation perturbée en raison d'un aléa naturel": hazard_env_protege,
        "Indicateur supplémentaire (sévérité 4)": add_env_protege,
        "Indicateur supplémentaire2 (sévérité 4)": add_2_env_protege,

    }

    # Classification de sévérité
    sev3 = {
        "description": "Enfants hors école qui ne subissent PAS de circonstances aggravantes ou enfants scolarisés dont l'éducation a été perturbée en raison de :",
        "ind1_scolarisés": prof_col,
        "ind2_scolarisés (aléa, sev3)": hazard_en_apprent,
        "ind3_scolarisés (supplémentaire, sev3)": add_en_apprent,
        "ind3_scolarisés (supplémentaire2, sev3)": add_2_en_apprent,

    }

    sev4 = {
        "description": "Enfants scolarisés dont l'éducation a été perturbée en raison de (indicateur scolarisé) ou enfants hors école confrontés aux circonstances aggravantes suivantes.",
        "indicateur_scolarisé (abri PDI)": idp_col,
        "indicateur_scolarisé (aléa, sev4)": hazard_env_protege,
        "indicateur_scolarisé (supplémentaire, sev4)": add_env_protege,
        "indicateur_scolarisé (supplémentaire2, sev4)": add_2_env_protege,

        "circonstances_aggravantes": st_session_state.get('selected_severity_4_barriers', []),
    }

    sev5 = {
        "description": "Enfants scolarisés dont l'éducation a été perturbée en raison de (indicateur scolarisé) ou enfants hors école confrontés aux circonstances aggravantes suivantes.",
        "indicateur_scolarisé": arme_col,
        "circonstances_aggravantes": st_session_state.get('selected_severity_5_barriers', []),
    }

    parameters = {
        "informations_generales": {
            "pays": pays,
            "date_du_calcul": datetime.now().strftime("%d/%m/%Y %H:%M"),
        },
        "indicateurs_msna_par_dimension": {
            "accès": acces_col,
            "conditions_d_apprentissage": bloc_apprentissage,
            "environnement_protégé": bloc_protege,
            "circonstances_aggravantes": barrier_col,
        },
        "classification_de_sévérité": {
            "niveau_de_sévérité_3": sev3,
            "niveau_de_sévérité_4": sev4,
            "niveau_de_sévérité_5": sev5,
        },
        "unité_administrative": {
            "unité_d_analyse": st_session_state.get('admin_var'),
            "décalage_admin": st_session_state.get('mismatch_admin', False),
        },
        "cycles_scolaires": {
            "tranches_d_age": st_session_state.get('vector_cycle'),
        }
    }
    return parameters

def generate_word_document_FR(parameters):
    """
    Génère le document Word FR des paramètres, en listant
    seulement les indicateurs réels (on saute 'no_indicator').
    """
    doc = docx.Document()
    doc.add_heading('Paramètres Utilisés comme Entrée pour le Calcul PiN', level=1)

    # Informations générales
    doc.add_heading('Informations Générales', level=2)
    general_info = parameters["informations_generales"]
    for key, value in general_info.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Indicateurs MSNA par dimension
    doc.add_heading('Indicateurs/Variables MSNA par Dimension', level=2)
    msna = parameters["indicateurs_msna_par_dimension"]
    for category, indicators in msna.items():
        if isinstance(indicators, dict):
            p = doc.add_paragraph(style='List Bullet')
            r = p.add_run(f"{category.replace('_', ' ').capitalize()}:")
            r.bold = True
            for description, indicator in indicators.items():
                # on affiche tout tel quel ici (même 'no_indicator'), pour rester fidèle à la version EN
                doc.add_paragraph(f"      {description}: {indicator}", style='List Bullet 2')
        else:
            p = doc.add_paragraph(style='List Bullet')
            r = p.add_run(f"{category.replace('_', ' ').capitalize()}: {indicators}")
            r.bold = True

    # Classification de sévérité
    doc.add_heading('Classification de Sévérité Utilisée pour ce Calcul', level=2)
    sev = parameters["classification_de_sévérité"]

    color_map = {
        "niveau_de_sévérité_3": RGBColor(255, 165, 0),  # orange clair
        "niveau_de_sévérité_4": RGBColor(255, 140, 0),  # orange foncé
        "niveau_de_sévérité_5": RGBColor(255, 69, 0),   # rouge-orangé
    }

    def list_inds(details, prefix="ind"):
        """Retourne la liste (clé, valeur) des champs indicateurs, en sautant 'no_indicator'."""
        items = []
        for k, v in details.items():
            if k.lower().startswith(prefix) and isinstance(v, str) and v != 'no_indicator':
                items.append((k, v))
        items.sort(key=lambda x: x[0])  # ordre stable
        return items

    for niveau, details in sev.items():
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(f"{niveau.replace('_', ' ').capitalize()}: ")
        r.bold = True
        if niveau in color_map:
            r.font.color.rgb = color_map[niveau]

        # description
        description = details.get("description", "")
        p.add_run(description + " ")

        # indicateurs (en gras, reliés par " et ")
        inds = list_inds(details, prefix="ind")
        if inds:
            for i, (_, val) in enumerate(inds):
                rr = p.add_run(val)
                rr.bold = True
                if i < len(inds) - 1:
                    p.add_run(" et ")
            p.add_run(".")

        # circonstances aggravantes (s'il y en a)
        if "circonstances_aggravantes" in details:
            for ex in details["circonstances_aggravantes"]:
                pp = doc.add_paragraph(style='List Bullet 2')
                pp.add_run(f"      {ex}")

    # Unité administrative
    doc.add_heading('Unité Administrative', level=2)
    admin = parameters["unité_administrative"]
    for key, value in admin.items():
        doc.add_paragraph(f"{key.replace('_', ' ').capitalize()}: {value}", style='List Bullet')

    # Cycles scolaires
    doc.add_heading('Cycles Scolaires', level=2)
    cycles = parameters.get("cycles_scolaires", {})
    tranches = cycles.get("tranches_d_age", [])
    doc.add_paragraph(f"Tranches d'âge: {tranches}", style='List Bullet')

    # Export
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output
