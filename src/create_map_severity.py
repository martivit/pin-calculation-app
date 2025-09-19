import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from functools import reduce
import geopandas as gpd
import matplotlib.pyplot as plt
import os, glob
from typing import Dict
import matplotlib.patches as mpatches
import re




def normalize_admin_columns(gdf, shapefile_path):
    fname = os.path.basename(shapefile_path).lower()
    colmap = {}
    if fname.startswith('afg_'):
        colmap = {
            'Prvnce_Cod': 'ADM2_PCODE',
            'Dstrct_Cod': 'ADM1_PCODE',
        }
    elif fname.startswith('drc_'):
        colmap = {
            'PROVINCE':   'ADM1_PCODE',
            'TERRITOIRE': 'ADM2_PCODE',
            'Pcode':      'ADM3_PCODE',
        }
    elif fname.startswith('car_'):
        colmap = {
            'admin1Pcod': 'ADM1_PCODE',
            'admin2Pcod': 'ADM2_PCODE',
            'admin3Pcod': 'ADM3_PCODE',
        }
    elif fname.startswith('syr_'):
        colmap = {
            'admin1Pcod': 'ADM1_PCODE',
            'admin2Pcod': 'ADM2_PCODE',
            'admin3Pcod': 'ADM3_PCODE',
        }
    # … add any other country‐specific cases here …

    # Apply the renames
    gdf.rename(columns=colmap, inplace=True)
    return gdf

def _strip_ner_prefix(series: pd.Series) -> pd.Series:
    """Convert 'NER...' → 'NE...' only if it appears as a prefix."""
    return (
        series.astype(str)
              .str.replace(r'^\s*NER', 'NE', regex=True)
              .str.strip()
              .replace({'nan': np.nan, 'None': np.nan})
    )







### map ########### map ########
desired_admin_level = {
    'MMR': 'ADM1_PCODE',
    'DRC': 'ADM3_PCODE',
}
DEFAULT_LEVEL = 'ADM2_PCODE'
shp_folder = "input_map"


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
color_mapping = {'light_beige', 'light_orange', 'dark_orange', 'darker_orange'}


def find_shapefile(shp_folder: str, country_code: str) -> str:
    """Return the first .shp in shp_folder whose basename (lowercased)
    starts with country_code.lower()."""
    code = country_code.lower()
    for path in glob.glob(os.path.join(shp_folder, "*.shp")):
        if os.path.basename(path).lower().startswith(code):
            return path
    raise FileNotFoundError(f"No .shp for '{country_code}' in {shp_folder}")

def make_map_severity(
    country: str,
    pin_data,               # pd.DataFrame
    shp_folder: str = "input_map",
    continuous_cmap: str = "Oranges",
    hpc_df: pd.DataFrame | None = None,      # NEW: DataFrame with HPC scope P-codes
    normalize_fn=normalize_admin_columns
) -> Dict[str, BytesIO]:
    """
    Returns PNG buffers keyed by field name for:
      - categorical severity (“Area severity” / “Sévérité de la zone”)
      - % severity level 5
      - % severity level 4
      - % severity level 3
      - % Tot PiN (3-5)

    All continuous maps get a true gradient colorbar.
    Missing areas are light gray, legends/colorbars sit to the right.
    """
    country_code = country.split('--')[-1].strip()
    shp_path = find_shapefile(shp_folder, country_code)
    gdf = gpd.read_file(shp_path)
    gdf = normalize_fn(gdf, shp_path)
     # --- FIX: repair invalid geometries BEFORE any dissolve/union ---
    try:
        from shapely import make_valid, set_precision  # Shapely 2.xhpc_df.rename
        # Make valid, keep polygonal, snap to a precision grid to avoid sliver conflicts
        gdf["geometry"] = make_valid(gdf.geometry)
        gdf = gdf[gdf.geometry.geom_type.isin(["Polygon", "MultiPolygon"])].copy()
        gdf["geometry"] = set_precision(gdf.geometry, grid_size=1e-8)
    except Exception:
        # Fallback for older Shapely: buffer(0) + keep polygonal only
        gdf["geometry"] = gdf.geometry.buffer(0)
        gdf = gdf[gdf.geometry.geom_type.isin(["Polygon", "MultiPolygon"])].copy()

    # One more pass: fix any remaining invalids individually
    bad = ~gdf.geometry.is_valid
    if bad.any():
        gdf.loc[bad, "geometry"] = gdf.loc[bad, "geometry"].buffer(0)

    # Final guard: drop empties and non-valids
    gdf = gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty & gdf.geometry.is_valid].copy()
    # --- END FIX ---
    # ---------- NIGER PCODE NORMALIZATION (NER → NE) ----------
    if country_code.upper() == 'NER':
        # fix all ADM*PCODE columns in the shapefile
        adm_pcode_cols = [c for c in gdf.columns if c.upper().startswith('ADM') and 'PCODE' in c.upper()]
        for c in adm_pcode_cols:
            gdf[c] = _strip_ner_prefix(gdf[c])

        # fix the first (key) column in the pin_data
        pin_data = pin_data.copy()
        pin_key_col = pin_data.columns[0]
        pin_data[pin_key_col] = _strip_ner_prefix(pin_data[pin_key_col])

        # fix HPC-scope DF if provided (assumes 2nd column holds the Pcodes, as per your comment)
        if hpc_df is not None and not hpc_df.empty:
            hpc_df = hpc_df.copy()
            if hpc_df.shape[1] >= 2:
                hpc_df.iloc[:, 1] = _strip_ner_prefix(hpc_df.iloc[:, 1])

    # 3) pick PIN‐code column & best ADM by string‐length
    pin_col = pin_data.columns[0]
    adm_cols = [c for c in gdf.columns if c.upper().startswith("ADM")]
    pin_len = pin_data[pin_col].astype(str).str.len().median()
    best_adm = min(adm_cols,
                   key=lambda c: abs(gdf[c].astype(str).str.len().median() - pin_len))

    # 4) merge
    merged = gdf.merge(pin_data, left_on=best_adm, right_on=pin_col, how="left")
    admin_level_gdf = (
        gdf
        .dissolve(by=best_adm, as_index=False)
        .set_index(best_adm)
    )
    plot_gdf = merged.copy()
    #plot_gdf.crs = None
    # HPC scope set
    hpc_set = set()
    #if hpc_df is not None:
        # assume second column holds the P-codes
        #hpc_set = set(hpc_df.iloc[:,1].astype(str))
    # ensure index type matches your pin_data key type
    #hpc_df = hpc_df.rename(columns={ hpc_df.columns[1]: best_adm })
    hpc_set = set()
    if hpc_df is not None and not hpc_df.empty and hpc_df.shape[1] >= 2:
        hpc_df = hpc_df.copy().rename(columns={hpc_df.columns[1]: best_adm})
        hpc_set = set(hpc_df[best_adm].astype(str))
    admin_level_gdf.index = admin_level_gdf.index.astype(str)
    # 5) colors for categorical map
    CAT_COLORS = {
        "1-2": "#FFF2CC",
        "3":   "#F4B183",
        "4":   "#ED7D31",
        "5":   "#C65911",
    }
    MISSING_COLOR = "#E5E4E2"
    MISSING_HPC  = "#ADD8E6"

    # 6) which continuous fields we expect
    CONT_FIELDS = {
        "% severity level 5", "% niveau de sévérité 5",
        "% severity level 4", "% niveau de sévérité 4",
        "% severity level 3", "% niveau de sévérité 3",
        "% Tot PiN (severity levels 3-5)", "% Tot PiN (niveaux de sévérité 3-5)"
    }

    # 7) specs: (title, possible cols, is_categorical)
    specs = [
        ("Area severity", ["Area severity", "Sévérité de la zone"], True),
    ]
    for fld in CONT_FIELDS:
        specs.append((fld, [fld], False))

    out: Dict[str, BytesIO] = {}

    for title, candidates, is_cat in specs:
        # pick the actual column
        field = next((c for c in candidates if c in plot_gdf.columns), None)
        if not field:
            continue

        fig, ax = plt.subplots(figsize=(8, 6))
        ax.set_aspect('equal')    # <— add this

        gdf.boundary.plot(ax=ax, edgecolor="#36454F", linewidth=0.1)

        # build masks
        missing = plot_gdf[field].isna()
        in_hpc  = plot_gdf[best_adm].astype(str).isin(hpc_set)

        if is_cat:
            #  — draw all areas grey first
            # 1) plot non-HPC missing
            mask1 = missing & ~in_hpc
            if mask1.any():
                plot_gdf[mask1].plot(facecolor=MISSING_COLOR, ax=ax, linewidth=0)
            # 2) plot HPC missing
            mask2 = missing & in_hpc
            if mask2.any():
                plot_gdf[mask2].plot(facecolor=MISSING_HPC, ax=ax, linewidth=0)
            #  — then overlay each severity class
            handles = [
                mpatches.Patch(color=MISSING_COLOR, label="Outside of HPC scope"),
                mpatches.Patch(color=MISSING_HPC,  label="No data in HPC scope")
            ]    
            for sev in ("1-2", "3", "4", "5"):
                sel = plot_gdf[plot_gdf[field].astype(str) == sev]
                if not sel.empty:
                    sel.plot(
                        facecolor=CAT_COLORS[sev],
                        edgecolor="black", linewidth=0.3,
                        ax=ax
                    )
                    handles.append(mpatches.Patch(
                        color=CAT_COLORS[sev], label=f"Severity {sev}"
                    ))
            ax.legend(
                handles=handles,
                title=title,
                loc="upper center", # "upper center"center left
                bbox_to_anchor=(1.02, 0.5),
                fontsize=8,             
                title_fontsize=10,      
                labelspacing=0.3       
            )

        else:
            # continuous: first fill missing
            plot_gdf.plot(
                column=field, cmap=continuous_cmap,
                edgecolor="none", linewidth=0,
                missing_kwds={"color":MISSING_COLOR},
                legend=False, ax=ax
            )
            # overplot HPC missing in blue
            mask2 = missing & in_hpc
            if mask2.any():
                plot_gdf[mask2].plot(facecolor=MISSING_HPC, ax=ax, linewidth=0)
            # colorbar
            sm = plt.cm.ScalarMappable(
                cmap=continuous_cmap,
                norm=plt.Normalize(vmin=plot_gdf[field].min(),
                                   vmax=plot_gdf[field].max())
            )
            sm._A = []
            cbar = fig.colorbar(sm, ax=ax, fraction=0.035, pad=0.04)
            cbar.set_label(title, rotation=270, labelpad=15)
            nos = [
                mpatches.Patch(color=MISSING_COLOR,
                               label="No data, outside HPC scope"),
                mpatches.Patch(color=MISSING_HPC,
                               label="No data, in HPC scope"),
            ]
            # place it below the axes
            ax.legend(
                handles=nos,
                loc="upper center",
                bbox_to_anchor=(0.5, -0.10),
                ncol=2,
                frameon=False
            )

        admin_level_gdf.boundary.plot( ax=ax, edgecolor="black", linewidth=0.5 )

        ax.set_axis_off()
        ax.set_title(f"{country}: {title}", fontsize=14)


        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
        buf.seek(0)
        out[title] = buf
        plt.close(fig)

    return out