# PiN Calculation App

A Streamlit-based web application for calculating People in Need (PiN) figures for education in humanitarian contexts, following the JIAF (Joint Intersectoral Analysis Framework) methodology.

## Table of Contents

- [Overview](#overview)
- [Application Structure](#application-structure)
  - [Main Application Files](#main-application-files)
  - [Pages Structure](#pages-structure)
    - [Page 1: Upload -- Education Data](#page-1-upload----education-data)
    - [Page 2: Calculation -- PiN](#page-2-calculation----pin)
    - [Page 3: Download -- PiN Figures and Other Outputs](#page-3-download----pin-figures-and-other-outputs)
  - [Source Functions (`src/`)](#source-functions-src)
    - [Core Calculation Functions](#core-calculation-functions)
    - [Output Generation Functions](#output-generation-functions)
    - [Hybrid Workflow Functions](#hybrid-workflow-functions)
  - [Supporting Files](#supporting-files)
    - [Translation Files](#translation-files)
    - [Data Directories](#data-directories)
    - [Debugging Scripts](#debugging-scripts)
- [Workflow Logic](#workflow-logic)
  - [Standard MSNA Countries](#standard-msna-countries)
  - [Hybrid Countries (Two-Step Process)](#hybrid-countries-two-step-process)
    - [Step 1: Initial Calculation](#step-1-initial-calculation)
    - [Step 2: Extrapolation & Final Calculation](#step-2-extrapolation--final-calculation)
  - [Alternative Data Sources (JENA/EMIS)](#alternative-data-sources-jenaemis)
  - [No OCHA Data Scenario](#no-ocha-data-scenario)
- [Installation and Setup](#installation-and-setup)
  - [Prerequisites](#prerequisites)
  - [Installation Steps](#installation-steps)
- [Usage](#usage)
  - [Basic Workflow](#basic-workflow)
  - [Data Requirements](#data-requirements)
  - [Configuration Tips](#configuration-tips)
- [Output Files](#output-files)
  - [Excel Files](#excel-files)
  - [Word Documents](#word-documents)
  - [Maps](#maps)
  - [GitHub Archive](#github-archive)
- [Debugging Outside Streamlit](#debugging-outside-streamlit)
  - [Why Debug Outside Streamlit?](#why-debug-outside-streamlit)
  - [Using Debug Scripts](#using-debug-scripts)
  - [Debugging Tips](#debugging-tips)
- [Contributing](#contributing)
  - [Development Priorities](#development-priorities)
- [License](#license)

## Overview

The PiN Calculation App automates the calculation of People in Need (PiN) figures for the Education sector in humanitarian crises. It processes Multi-Sectoral Needs Assessment (MSNA) data, applies JIAF severity scoring, and generates formatted Excel outputs with snapshots and visualizations.

**Key Features:**
- Multi-language support (English and French)
- Multiple data source support (MSNA, JENA, EMIS)
- Hybrid calculation workflows for specific countries
- Automated severity scoring based on JIAF framework
- Geographic visualization with severity maps
- GitHub integration for automatic file archiving
- PDF snapshot generation for HNO (Humanitarian Needs Overview) submissions

## Application Structure

### Main Application Files

```
PiN_app_dev/
├── GEC_PiN.py                    # Main Streamlit app entry point
├── shared_utils.py                # Shared utilities (language selector, etc.)
├── requirements.txt               # Python dependencies
└── .streamlit/
    └── secrets.toml              # GitHub tokens and credentials (not in repo)
```

**`GEC_PiN.py`**: Main entry point that initializes the Streamlit app, sets up page configuration, handles authentication, and provides the landing page with navigation to the three calculation pages.

**`shared_utils.py`**: Contains shared functions used across all pages, primarily the `language_selector()` function that manages multilingual support.

### Pages Structure

The application has three sequential pages that guide users through the PiN calculation process:

```
pages/
├── 1_📁_Upload_--_Education_Data.py
├── 2_📊_Calculation_--_PiN.py
└── 3_📋_Download_--_PiN_figures_and_other_outputs.py
```

#### **Page 1: Upload -- Education Data**
- **Purpose**: Data upload and initial configuration
- **User Actions**:
  - Select country and education cycle parameters
  - Upload MSNA dataset (household and education data)
  - Upload KoBoToolbox form (survey structure)
  - Upload OCHA population figures
  - For alternative scenarios: upload JENA or EMIS data
- **Key Functions Called**:
  - Validates uploaded files
  - Stores data in `st.session_state` for use in subsequent pages
  - Performs initial data quality checks

#### **Page 2: Calculation -- PiN**
- **Purpose**: Configure indicators and severity thresholds
- **User Actions**:
  - Map dataset columns to PiN indicators (access, disruption, barriers)
  - Define severity thresholds for barriers (what constitutes severity 4 vs 5)
  - Configure population group mappings (host, IDP, returnee, refugee)
  - Select administrative boundary variables
  - Enable/disable additional custom indicators
- **Key Functions Called**:
  - Column mapping validation
  - Severity threshold configuration
  - Preview of selected indicators
- **Output**: Configuration stored in `st.session_state` for Page 3

#### **Page 3: Download -- PiN Figures and Other Outputs**
- **Purpose**: Execute calculations and generate outputs
- **Main Process Flow**:
  1. Clean and prepare dataset → `clean_make_dataset()`
  2. Add severity scores to indicators → `add_severity()`
  3. Calculate PiN figures → `calculatePIN()` (or variant)
  4. Generate outputs → `create_output()`, `create_snapshot_PiN()`
  5. Create downloadable ZIP file
  6. Upload results to GitHub repository
- **Key Functions Called**: See [Source Functions](#source-functions-src) section
- **Output**: ZIP file containing Excel results, Word snapshots, parameter documents, and severity maps

### Source Functions (`src/`)

The `src/` folder contains all calculation logic and output generation functions:

```
src/
├── clean_dataset.py                              # Data cleaning and preparation
├── add_PiN_severity.py                          # Severity scoring for indicators
├── calculation_for_PiN_Dimension.py             # Main PiN calculation (standard)
├── calculation_for_PiN_Dimension_with_JENA.py   # PiN calculation with JENA data
├── calculation_for_PiN_Dimension_with_EMIS.py   # PiN calculation with EMIS data
├── calculation_for_PiN_Dimension_NO_OCHA.py     # PiN calculation without OCHA data
├── calculation_for_PiN_Dimension_NO_OCHA_2025.py # Updated version for 2025
├── update_re_calculation_for_PiN.py             # Recalculation for hybrid step 2
├── vizualize_PiN.py                             # Excel output formatting
├── snapshot_PiN.py                              # Word snapshot generation (EN)
├── snapshot_PiN_FR.py                           # Word snapshot generation (FR)
├── save_parameter.py                            # Parameter documentation (EN)
├── save_parameter_FR.py                         # Parameter documentation (FR)
├── create_map_severity.py                       # Geographic visualization
├── make_output1_platform.py                     # Hybrid step 1 output merger
├── create_output1_excel.py                      # Hybrid step 1 Excel formatting
├── extrapolation.py                             # Delta method extrapolation
└── cases_country_helpers.py                     # Country-specific helper functions
```

#### **Core Calculation Functions**

**`clean_make_dataset()`** (`clean_dataset.py`)
- Cleans and standardizes the uploaded MSNA data
- Creates age groups (ECE, primary, upper primary, secondary)
- Normalizes population status groups
- Handles missing values and data type conversions
- Returns cleaned education and household datasets with processing messages

**`add_severity()`** (`add_PiN_severity.py`)
- Maps indicator responses to JIAF severity scores (1-5)
- Applies severity thresholds for barriers
- Handles special cases for disruption indicators
- Creates severity columns for: access, teacher disruption, IDP disruption, armed group disruption, natural hazards, barriers
- Returns dataset with added severity columns

**`calculatePIN()`** (`calculation_for_PiN_Dimension.py`)
- **Main PiN calculation function** for standard MSNA countries
- Aggregates severity scores by dimension (Access, Learning Environment, Protected Environment)
- Calculates maximum severity across dimensions
- Disaggregates by administrative area, population group, age group, and gender
- Merges with OCHA population figures to calculate absolute PiN numbers
- Returns multiple outputs: severity tables, dimension tables, indicator breakdowns, totals

**Variant Calculation Functions:**
- `calculatePIN_with_JENA()`: Integrates JENA (Joint Education Needs Assessment) data
- `calculatePIN_with_EMIS()`: Integrates EMIS (Education Management Information System) data
- `calculatePIN_NO_OCHA()` / `calculatePIN_NO_OCHA_2025()`: Calculates PiN percentages without OCHA population data
- `UPDATE_calculatePIN()`: Used in hybrid step 2 to recalculate after extrapolation

#### **Output Generation Functions**

**`create_output()`** (`vizualize_PiN.py`)
- Creates the main Excel output file with multiple sheets:
  - PiN TOTAL: Overall PiN by severity
  - By administrative area
  - By population group
  - By age group and gender
  - Dimension breakdowns
- Applies formatting, color coding, and conditional formatting
- Includes parameter documentation sheet

**`create_indicator_output()`** (`vizualize_PiN.py`)
- Generates detailed Excel file showing PiN breakdown by each indicator
- Useful for understanding which indicators drive severity scores

**`create_snapshot_PiN()` / `create_snapshot_PiN_FR()`** (`snapshot_PiN.py` / `snapshot_PiN_FR.py`)
- Creates Word document snapshots with:
  - Key PiN figures summary
  - Severity breakdown charts
  - Gender disaggregation
  - Population group comparison
- Formatted for HNO submission
- Available in English and French

**`generate_word_document()` / `generate_word_document_FR()`** (`save_parameter.py` / `save_parameter_FR.py`)
- Documents all parameters and configuration choices used in the calculation
- Records: country, education cycle, indicator mappings, severity thresholds, data sources
- Critical for reproducibility and transparency

**`make_map_severity()`** (`create_map_severity.py`)
- Generates geographic visualizations showing PiN severity by administrative area
- Creates PNG images for inclusion in reports
- Color-coded by severity level (1-5)

#### **Hybrid Workflow Functions**

**`merge_2025_contextDB()`** (`make_output1_platform.py`)
- **Hybrid Step 1**: Merges calculated PiN from MSNA-covered areas with secondary data
- Integrates: ACLED conflict data, Information Insight (II) indicators, empty columns for manual filling
- Creates temporary output file for users to complete

**`create_output1_user()`** (`create_output1_excel.py`)
- Formats the hybrid step 1 output with color-coding and clear instructions
- Highlights cells that need manual completion

**`extrapolate_df_2025_updated()`** (`extrapolation.py`)
- **Hybrid Step 2**: Uses delta method to extrapolate PiN to non-covered areas
- Compares 2024 vs 2025 patterns in covered areas
- Applies trends to uncovered areas
- Returns extrapolated PiN figures for entire country

### Supporting Files

#### **Translation Files**
```
en.json       # English translations for all UI text
fr.json       # French translations for all UI text
```

These JSON files contain all user-facing text, allowing complete interface translation between English and French.

#### **Data Directories**
```
input/              # User-uploaded data files
input_map/          # Geographic boundary files for mapping
context_DB/         # Secondary data (ACLED, II) for hybrid countries
pin2024_cat/        # 2024 PiN data for extrapolation
platform_PiN_output/  # Archived outputs (uploaded to GitHub)
pics/               # Logo and UI images
icon/               # App icon
output_validation/  # Quality check outputs (if enabled)
```

#### **Debugging Scripts**

**Outside-Streamlit Execution Files:**
```
run_PiNcalculation_outside_streamlit.py        # Debug standard MSNA workflow
run_PiNcalculation_outside_streamlit2.py       # Debug hybrid workflow
run_PiNcalculation_outside_streamlit_EMIS.py   # Debug EMIS workflow
run_PiNcalculation_outside_streamlit_JENA.py   # Debug JENA workflow
```

These scripts allow developers to:
- Run calculations without launching Streamlit
- Test individual functions in isolation
- Debug with IDE tools and breakpoints
- Rapidly iterate on calculation logic
- Use hardcoded file paths instead of uploads

**Usage Example:**
```python
# In run_PiNcalculation_outside_streamlit.py
country = "Somalia -- SOM"
dataset_path = "input/Somalia_MSNA_2025.xlsx"
# ... set all parameters directly
# Run calculations
results = calculatePIN(country, edu_data, household_data, ...)
```

## Workflow Logic

### Standard MSNA Countries

**Countries**: Most humanitarian contexts with comprehensive MSNA coverage

**Flow**:
1. **Upload** (Page 1): MSNA dataset + OCHA figures
2. **Configure** (Page 2): Map indicators, set thresholds
3. **Calculate** (Page 3):
   - Clean data → `clean_make_dataset()`
   - Add severity → `add_severity()`
   - Calculate PiN → `calculatePIN()`
   - Generate outputs → `create_output()`, `create_snapshot_PiN()`
4. **Download**: Single ZIP file with all results

**Outputs**:
- `PiN_results_{country}_{timestamp}.xlsx`
- `PiN_by_indicator_{country}_{timestamp}.xlsx`
- `PiN_snapshot_{country}_{timestamp}.docx`
- `Parameters_Input_Document_{timestamp}.docx`
- Severity maps (PNG files)

### Hybrid Countries (Two-Step Process)

**Countries**: CAR, Ethiopia, DRC, Lebanon, Somalia, South Sudan

**Why Hybrid?**: MSNA only covers targeted areas; need to extrapolate to entire country

#### **Step 1: Initial Calculation**

1. **Upload** (Page 1): MSNA for covered areas + OCHA figures
2. **Configure** (Page 2): Standard indicator mapping
3. **Calculate** (Page 3):
   - Calculate PiN for covered areas → `calculatePIN()`
   - Merge with secondary data → `merge_2025_contextDB()`
   - Format for user completion → `create_output1_user()`
4. **Download**: Temporary file to fill
   - `PiN_temporary_to_fill_{country}_{timestamp}.xlsx`
   - User manually completes empty cells using secondary data

#### **Step 2: Extrapolation & Final Calculation**

1. **Upload** (Page 1): 
   - Completed temporary file from Step 1
   - 2024 PiN data for comparison
2. **Skip** Page 2 (configuration already done)
3. **Calculate** (Page 3):
   - Extrapolate using delta method → `extrapolate_df_2025_updated()`
   - Recalculate country-wide PiN → `UPDATE_calculatePIN()`
   - Generate final outputs
4. **Download**: Final PiN results for entire country

**Key Difference**: Step 2 uses `st.session_state.step_2_hpc = True` flag to trigger extrapolation workflow

### Alternative Data Sources (JENA/EMIS)

**Countries**: Niger, Nigeria, Mozambique (when `data_combination != 'mmmm'`)

**JENA (Joint Education Needs Assessment)**:
- School-level assessment data
- Integrates with MSNA using `calculatePIN_with_JENA()`
- Combines household-level (MSNA) and school-level (JENA) indicators

**EMIS (Education Management Information System)**:
- Government education system data
- Enrollment figures, school infrastructure data
- Integrates with MSNA using `calculatePIN_with_EMIS()`
- Helps validate and complement MSNA findings

**Data Combination Codes**:
- `'mjjm'`: MSNA + JENA (uses JENA data)
- `'emmm'`, `'eemm'`, `'eeem'`: MSNA + EMIS (uses EMIS data)
- `'mmmm'`: MSNA only (standard workflow)

### No OCHA Data Scenario

**When**: User indicates OCHA population figures are not available

**Limitation**: Cannot calculate absolute PiN numbers, only percentages

**Flow**:
1. Calculate severity percentages → `calculatePIN_NO_OCHA_2025()`
2. Generate indicator breakdown → `create_indicator_output_no_ocha()`
3. Generate percentage output → `create_pin_raw_output()`

**Outputs**:
- `PiN_percentage_{country}_{timestamp}.xlsx` (severity %, not absolute numbers)
- `PiN_by_indicator_{country}_{timestamp}.xlsx`
- `Parameters_Input_Document_{timestamp}.docx`

## Installation and Setup

### Prerequisites

- Python 3.8 or higher
- Git (for cloning repository)

### Installation Steps

1. **Clone the repository**:
```bash
git clone https://github.com/Global-Education-Cluster-PiN/pin-calculation-app.git
cd pin-calculation-app
```

2. **Create virtual environment**:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies**:
```bash
pip install -r requirements.txt
```

4. **Configure secrets** (for GitHub integration):
Create `.streamlit/secrets.toml`:
```toml
[github]
token = "your_github_personal_access_token"
```

5. **Run the application**:
```bash
streamlit run GEC_PiN.py
```

The app will open in your browser at `http://localhost:8501`

## Usage

### Basic Workflow

1. **Launch Application**: Run `streamlit run GEC_PiN.py`
2. **Select Language**: Choose English or French
3. **Navigate to Page 1**: Upload your data files
4. **Navigate to Page 2**: Configure indicators and thresholds
5. **Navigate to Page 3**: Execute calculation and download results

### Data Requirements

**Required Files**:
- MSNA dataset (Excel): Household and education loop data
- KoBoToolbox form (Excel): Survey structure with question types
- OCHA population figures (Excel): Population by admin area and group

**Optional Files** (scenario-dependent):
- JENA data (for JENA integration)
- EMIS data (for EMIS integration)
- 2024 PiN data (for hybrid step 2)
- Completed temporary file (for hybrid step 2)

### Configuration Tips

**Indicator Mapping**:
- Access: Use "currently attending school" type questions
- Disruption: Use "reasons for missing school" or "disruption experienced"
- Barriers: Use "main barrier to education" type questions

**Severity Thresholds**:
- Severity 4: Moderate barriers (e.g., "lack of school supplies")
- Severity 5: Critical barriers (e.g., "school too dangerous", "forced marriage")

**Population Groups**:
- Ensure consistent naming between MSNA and OCHA data
- Use mapping feature if names don't match exactly

## Output Files

### Excel Files

**`PiN_results_{country}_{timestamp}.xlsx`**:
- Main output file with multiple sheets
- PiN by severity, admin area, population group, age, gender
- Formatted with color-coding and conditional formatting

**`PiN_by_indicator_{country}_{timestamp}.xlsx`**:
- Detailed breakdown showing which indicators contribute to each severity level
- Useful for programmatic targeting decisions

**`PiN_temporary_to_fill_{country}_{timestamp}.xlsx`** (Hybrid only):
- Partially completed file requiring manual input
- Color-coded cells show what needs completion

### Word Documents

**`PiN_snapshot_{country}_{timestamp}.docx`**:
- Summary document for HNO submission
- Key figures, charts, and breakdowns
- Ready for inclusion in reports

**`Parameters_Input_Document_{timestamp}.docx`**:
- Complete record of all calculation parameters
- Critical for transparency and reproducibility

### Maps

**`{country}_{admin_level}_severity.png`**:
- Geographic visualization of PiN severity
- Color-coded by severity level
- One map per administrative level

### GitHub Archive

All outputs are automatically uploaded to the GitHub repository at:
```
platform_PiN_output/{country}/
```

This creates a historical archive of all calculations for audit and comparison purposes.

## Debugging Outside Streamlit

### Why Debug Outside Streamlit?

- **Faster iteration**: No need to reload entire Streamlit app
- **Better debugging tools**: Use IDE breakpoints and inspectors
- **Easier testing**: Test individual functions in isolation
- **Reproducible scenarios**: Hardcode parameters for consistent testing

### Using Debug Scripts

1. **Choose the appropriate script** based on your scenario:
   - Standard: `run_PiNcalculation_outside_streamlit.py`
   - Hybrid: `run_PiNcalculation_outside_streamlit2.py`
   - JENA: `run_PiNcalculation_outside_streamlit_JENA.py`
   - EMIS: `run_PiNcalculation_outside_streamlit_EMIS.py`

2. **Modify file paths** in the script:
```python
# Example from run_PiNcalculation_outside_streamlit.py
country = "Somalia -- SOM"
dataset_path = "input/Somalia_MSNA_2025.xlsx"
kobo_path = "input/Somalia_kobo_form.xlsx"
ocha_path = "input/Somalia_OCHA_figures.xlsx"
```

3. **Set parameters** directly:
```python
access_var = "edu_access"
teacher_disruption_var = "edu_disruption_teacher"
# ... etc
```

4. **Run the script**:
```bash
python run_PiNcalculation_outside_streamlit.py
```

5. **Inspect outputs**: Results are saved to `output/` directory

### Debugging Tips

- Use `print()` statements liberally to trace data flow
- Add `import pdb; pdb.set_trace()` for interactive debugging
- Check intermediate dataframes with `df.head()`, `df.info()`
- Validate severity scoring by checking severity column distributions
- Compare outputs against known good results from Streamlit runs

## Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/improvement`
3. **Follow coding standards**:
   - Use descriptive variable names
   - Add docstrings to all functions
   - Comment complex logic
   - Follow PEP 8 style guide
4. **Test thoroughly**:
   - Test with multiple countries
   - Test edge cases (missing data, etc.)
   - Verify outputs match expected results
5. **Submit a pull request** with clear description of changes

### Development Priorities

- [ ] Add automated testing suite
- [ ] Improve error handling and user feedback
- [ ] Optimize calculation performance for large datasets
- [ ] Expand geographic visualization options
- [ ] Add data quality check dashboards

## License

This project is developed by the Global Education Cluster for humanitarian use. 

For questions or support, contact: [global.education.cluster@humanitarianresponse.info]

---

**Version**: 2.0 (2025 HNO Cycle)  
**Last Updated**: January 2026  
**Maintained by**: Global Education Cluster Technical Team