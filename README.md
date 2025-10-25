# HR Personnel & CV Generator

A Streamlit-based web application for managing key personnel data and generating bulk CVs for HR tender submissions.

## Overview

This application streamlines the process of managing personnel data, searching for qualified candidates based on roles, and automatically generating professional CVs in bulk for tender documents. It features an intuitive 3-step workflow with advanced search capabilities and inline editing.

## Features

### 🔄 3-Step Workflow

1. **Data Preparation & Validation**
   - Auto-loads personnel and project data from system files
   - Upload custom personnel Excel files
   - Real-time data validation and format analysis
   - Automatic date format conversion (supports DD-MM-YYYY, MM-YYYY, year-only formats)

2. **Role Definition & Search (Optional)**
   - Define multiple roles with custom criteria
   - Advanced search filters:
     - Keyword-based qualification matching (OR logic)
     - Exact word match or contains mode
     - Minimum years of experience
     - Diploma inclusion/exclusion toggle
   - Detailed categorization of search results:
     - Fully matched candidates
     - Title + Qualification match (insufficient experience)
     - Qualification + Experience match (title mismatch)
     - Qualification only matches
   - Comprehensive summary reports by role

3. **Edit & Generate**
   - **Job Title Mode Selection**: Choose between using existing job titles or assigning new roles
   - Define custom roles for assignment (when using "Assign New Roles" mode)
   - Inline table editing with auto-save
   - Add individual users via form with role dropdown
   - Bulk assignment tools (Job Title with role dropdown, From Date, Qualification)
   - Auto-calculation of Years of Experience
   - Delete selected rows
   - Generate bulk CVs as Word document
   - Download edited Excel data

### 📊 Smart Data Management

- **Auto-Save**: All changes automatically saved to temp files
- **Format Conversion**: Automatically converts various date formats to MM-YYYY
- **YOE Calculation**: Automatically calculates Years of Experience as integer (floor value)
- **Empty Row Removal**: Automatically removes empty rows
- **Data Validation**: 
  - Real-time format validation while editing (non-blocking)
  - Pre-download/generate comprehensive validation
  - **Critical validations (blocking)**:
    - Missing Names, Job Titles, Qualifications, From Dates
  - **Warning validations (non-blocking)**:
    - Invalid date formats, Zero years of experience
  - Detailed validation reports with affected row details (first 10 shown)
  - Clear separation between critical issues and warnings
- **Date Handling**: "To" date always set to "Present"
- **Assignment Tracking**: Live metrics showing job title assignment status

### 🔍 Advanced Search Capabilities

- **Flexible Qualification Filtering**: Search by keywords with OR logic
- **Diploma Control**: Include or exclude diploma holders
- **Experience Requirements**: Set minimum years of experience
- **Multi-Category Results**: Candidates categorized by match quality
- **Summary Analytics**: Visual metrics showing matches vs. requirements

### 📝 CV Generation

- **Bulk Processing**: Generate CVs for all personnel in one click
- **Professional Format**: Times New Roman, 8pt font, proper spacing
- **Smart Project Assignment**: Randomly assigns eligible projects based on timeline overlap
- **Structured Layout**: Two-table format with personal info and experience details
- **Company Details**: Pre-filled with Pioneer Foundation Engineers Pvt. Ltd. information

## Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Setup

1. Clone or download this repository

2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure the following directory structure exists:
   ```
   project_root/
   ├── app2.py
   ├── requirements.txt
   ├── input_csv/
   │   └── personnel.xlsx
   ├── input_excel/
   │   └── exployee.xlsx
   └── temp_uploads/
   ```

## Usage

### Starting the Application

Run the Streamlit app:
```bash
streamlit run app2.py
```

The application will open in your default web browser at `http://localhost:8501`

### Workflow Guide

#### Step 1: Load & Review Data

1. **System Files (Auto-loaded)**
   - Personnel file: `input_csv/personnel.xlsx`
   - Project info: `input_excel/exployee.xlsx` (sheet: `project_info`)
   - Click "🔄 Reload Personnel" if files are updated

2. **Upload Custom Files**
   - Check "📤 Upload Custom Files" to upload your own personnel data
   - Select the appropriate sheet from your Excel file
   - Project info always loads from system file

3. **Confirm Data**
   - Review data validation results
   - Click "✔️ Confirm & Search Personnel" to proceed with search
   - Or click "⏩ Skip Search, Go to Edit" to skip search and directly edit

#### Step 2: Search Personnel (Optional)

1. **Define Roles**
   - Enter Role Name (e.g., "Civil Engineer")
   - Set Required Count
   - Add Qualification Keywords (comma-separated, uses OR logic)
   - Choose search mode: "Contains" or "Exact Word Match"
   - Set diploma inclusion preference
   - Set minimum years of experience
   - Click "💾 Save Role"

2. **Run Search**
   - Click "🔎 Start Search" to execute
   - Review categorized results:
     - ✅ Fully Matched
     - ⚠️ Title+Qual (Low Exp)
     - 🔍 Qual+Exp (Job Title Mismatch)
     - 📋 Qualification Only
     - ❌ No Match
   - Check summary metrics

3. **Multiple Roles**
   - Add as many roles as needed
   - Search processes all roles simultaneously
   - Click "🗑️ Clear All Roles" to reset

#### Step 3: Edit & Generate

1. **Job Title Mode Selection** (First time in Step 3)
   - **📋 Use Existing Job Titles**: Keep current job titles from Excel file
     - Edit job titles freely as text
     - Works like traditional editing
   
   - **🎯 Assign New Roles**: Define roles and assign them to personnel
     - Clear all existing job titles
     - Define at least one role before proceeding
     - Add multiple roles as needed
     - Job Title becomes a dropdown with defined roles
     - "Custom" option available for exceptions

2. **Define Roles** (If "Assign New Roles" mode selected)
   - Enter role names (e.g., "Civil Engineer", "Project Manager")
   - Add multiple roles to the list
   - Remove individual roles if needed
   - Clear all roles and start over
   - Must confirm at least one role before proceeding to edit

3. **Inline Editing**
   - Edit Name, Qualification, Job Title, From (MM-YYYY) directly in table
   - **Job Title**: Text input (existing mode) or Dropdown (assign roles mode)
   - To date auto-set to "Present"
   - Years of Experience auto-calculated
   - Click + button to add new rows
   - Empty rows automatically removed on save

4. **Add Single User**
   - Expand "➕ Add Single User" section
   - Fill all mandatory fields
   - From date must be MM-YYYY format (e.g., 01-2020)
   - YOE calculated automatically
   - Click "✅ Add User"

5. **Bulk Operations**
   - Select rows using checkboxes
   - Choose field to update: Job Title, From, or Qualification
   - **Job Title Assignment**:
     - **Assign Roles Mode**: Select from defined roles dropdown + "Custom" option
     - **Existing Mode**: Enter job title as text
   - **From Date**: Enter date in MM-YYYY format
   - **Qualification**: Enter qualification as text
   - Click "✅ Apply to Selected"
   - Delete selected rows with "🗑️ Delete Selected"

6. **Data Validation** (Automatic before download/generate)
   - **Critical Errors (Must Fix) - BLOCKS Download & CV Generation**:
     - ❌ Missing Names: All personnel must have names
     - ❌ Missing Job Titles: All personnel must have job titles assigned
     - ❌ Missing Qualifications: All personnel must have qualifications
     - ❌ Missing From Dates: All personnel must have start dates
   - **Warnings (Recommended to Fix) - Does NOT block**:
     - ⚠️ Invalid From Date Format: Dates not in MM-YYYY format
     - ⚠️ Zero Years of Experience: Personnel with 0 YOE (may need correction)
   - **Validation Report Features**:
     - Summary showing count of critical issues and warnings
     - Detailed list of affected rows (shows first 10, indicates if more)
     - Row identifiers with personnel names for easy location
     - Clear explanations of each issue
     - Collapsible expander for easy access

7. **Download & Generate**
   - **Download Excel**: Get current personnel data as Excel file
     - Blocked if critical validation errors exist
     - Shows error message with details
   - **Generate CVs**: Create Word document with all CVs
     - Blocked if critical validation errors exist
     - CVs include personal info and experience
     - Projects randomly assigned based on timeline overlap
     - Download as `Employees_CV.docx`

## File Structure

### Input Files

#### Personnel File (`input_csv/personnel.xlsx`)
Required columns:
- **Name**: Employee full name
- **Qualification**: Educational qualification (e.g., "B.E. Civil", "Diploma Civil")
- **Job Title**: Current position/role
- **From**: Start date (supports MM-YYYY, DD-MM-YYYY, year-only)
- **To**: End date (auto-set to "Present")
- **Years of Experience**: Auto-calculated from "From" date

#### Project Info File (`input_excel/exployee.xlsx`, sheet: `project_info`)
Required columns:
- **Start Date**: Project start date
- **Work Completion date**: Project completion date
- **Company / Project / Position**: Project description
- **Relevant Technical & Managerial Experience**: Detailed experience points (use hyphen-separated points)

### Output Files

- **Employees_CV.docx**: Generated CVs in Word format
- **temp_uploads/personnel_temp_YYYYMMDD_HHMMSS.xlsx**: Auto-saved temp files
- **Personnel_download.xlsx**: Manual Excel download

## Data Formats

### Date Formats

**Supported Input Formats:**
- MM-YYYY: `01-2020`, `11-2022`
- DD-MM-YYYY: `01-01-2006`, `15-06-2022`
- Year only: `2017`, `2020`
- MM/YYYY: `01/2020`, `11/2022`
- Excel datetime objects (auto-converted)

**Output Format:**
- Always MM-YYYY: `01-2020`, `06-2022`
- "Present" for current employment

### Years of Experience Calculation

- Calculated as integer years (floor value, no decimals)
- Based on "From" date to current date
- Formula: `floor((current_month - from_month) / 12)`
- Example: 2.8 years → 2 years

### Qualification Keywords

**Search Examples:**
- Single keyword: `civil` (matches "B.E. Civil", "Civil Engineering")
- Multiple keywords (OR logic): `civil, mechanical` (matches EITHER civil OR mechanical)
- Exact word match: Matches complete words only
- Contains mode: Matches keywords anywhere in qualification

**Diploma Handling:**
- Include Diploma: Matches all qualifications including diploma
- Exclude Diploma: Only degree holders (filters out "Diploma" qualifications)

## Configuration

### File Paths
```python
PERSONNEL_PATH = "input_csv/personnel.xlsx"
PROJECT_WB_PATH = "input_excel/exployee.xlsx"
PROJECT_INFO_SHEET = "project_info"
SAVE_DIR = "temp_uploads"
OUTPUT_DOCX = "Employees_CV.docx"
```

### CV Formatting
```python
FONT_NAME = "Times New Roman"
FONT_SIZE = 8
LINE_SPACING = 1.15
INDENT_CM = 0.12
```

### Company Details (Pre-filled in CVs)
- **Name**: Pioneer Foundation Engineers Private Limited
- **Address**: Boomerang, B-2, 508/509, Off Chandivali Farm Rd, Chandivali, Powai, Mumbai, Maharashtra 400072
- **Telephone**: 022 4801 1311
- **Mobile**: +91 99209 03578
- **Email**: sales@pfepl.com

## Features Details

### Job Title Modes

The application offers two distinct modes for managing job titles in Step 3:

#### 📋 Use Existing Job Titles Mode
- **Purpose**: Keep and edit existing job titles from the loaded Excel file
- **Behavior**:
  - Job titles remain as they are in the source file
  - Edit job titles as free text (no restrictions)
  - Bulk assignment uses text input
  - Add Single User form uses text input
  - Best for: Minor edits, retaining current structure

#### 🎯 Assign New Roles Mode
- **Purpose**: Standardize job titles by assigning from a predefined list of roles
- **Behavior**:
  - All existing job titles are cleared when this mode is selected
  - Define at least one role before proceeding (can add multiple)
  - Job Title column becomes a dropdown in the table
  - Bulk assignment shows role dropdown with "Custom" option
  - Add Single User form shows role dropdown with "Custom" option
  - Best for: Standardization, ensuring consistency, tender requirements

#### Switching Modes
- Mode can be changed at any time using "🔄 Change Mode" button
- **Warning**: Switching from "Assign New Roles" back to "Use Existing" will retain the assigned roles as text
- Switching from "Use Existing" to "Assign New Roles" will clear all job titles
- Defined roles are preserved when switching modes

#### Custom Option
- When in "Assign New Roles" mode, "Custom" option is always available
- Select "Custom" to enter a job title that's not in the defined roles list
- Useful for exceptions or one-off positions
- Custom entries are saved as-is (not added to defined roles list)

### Auto-Save System
- First edit creates timestamped file in `temp_uploads/`
- Subsequent edits overwrite the same file
- Saves after every change:
  - Table edits
  - User additions
  - Bulk assignments
  - Row deletions

### Search Logic
- **Role Name**: Matches against "Job Title" (case-insensitive, contains)
- **Qualification Keywords**: OR logic (matches ANY keyword, not all)
- **Experience Filter**: Minimum threshold (greater than or equal to)
- **Diploma Filter**: Include/exclude based on "Diploma" keyword in qualifications

### CV Generation Algorithm
1. For each employee:
   - Parse employee's From/To dates
   - Find all projects that overlap with employee tenure
   - Randomly select ONE eligible project
   - Avoid duplicate project assignments (when possible)
   - Use employee's From/To dates in CV (not project dates)
   - Display "Present" for To date
   - Include project description as bullet points

## Tips & Best Practices

1. **Job Title Mode**: Choose "Assign New Roles" for standardization, "Use Existing" for flexibility
2. **Define Roles Early**: If using "Assign New Roles", define all roles before editing personnel
3. **Custom Option**: Use "Custom" sparingly - defeats the purpose of standardization
4. **Date Entry**: Always use MM-YYYY format (e.g., 01-2020) for consistency
5. **Bulk Operations**: Select multiple rows for faster editing, especially useful with role dropdowns
6. **Search Strategy**: Start with broad criteria, then refine
7. **Qualification Keywords**: Use comma-separated list for OR logic
8. **Auto-Save**: Changes save automatically, but download Excel for backup
9. **Empty Rows**: Don't worry about empty rows, they're auto-removed
10. **Format Warnings**: Non-blocking, you can save and fix later
11. **Project Assignment**: Random but timeline-aware (ensures logical CVs)
12. **Mode Switching**: Be cautious when switching modes - job titles may be cleared

## Troubleshooting

### Files Not Loading
- Check file paths in configuration
- Ensure Excel files are not open in another program
- Verify sheet names match configuration

### Date Format Issues
- Use MM-YYYY format: `01-2020`, `11-2022`
- Avoid spaces, special characters
- Year must be 4 digits

### Search Returns No Results
- Check if personnel data is loaded
- Verify keyword spelling
- Try "Contains" mode instead of "Exact Word"
- Lower minimum experience requirement

### Cannot Download or Generate CVs
- **Check validation results**: Click on "🔍 Data Validation Results" expander
- **Critical Issues (Must Fix)**:
  - **Missing Names**: Add names to all rows with empty Name field
  - **Missing Job Titles**: Assign job titles to all personnel
    - In "Use Existing" mode: Type job title directly
    - In "Assign New Roles" mode: Select from dropdown or choose "Custom"
    - Use bulk assignment tool for multiple rows at once
  - **Missing Qualifications**: Add qualifications to all rows
  - **Missing From Dates**: Add start dates in MM-YYYY format (e.g., 01-2020)
- **How to Find Issues**:
  - Validation report shows affected row numbers and names
  - Shows first 10 affected rows, indicates if more exist
  - Look for warning banner above table with missing field counts
  - Check metrics: "Job Titles Assigned" shows X/Total
- **Quick Fix Tips**:
  - Select multiple rows → Use bulk assignment
  - Add missing data directly in table
  - Delete completely empty rows if not needed

### CV Generation Fails
- Ensure project info file is loaded
- Check project date columns exist
- Verify at least some personnel data exists
- Ensure all job titles are assigned (check validation)

## License

This project is proprietary software for Pioneer Foundation Engineers Private Limited.

## Support

For technical support or questions, please contact the development team.
