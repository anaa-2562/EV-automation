# Audentes Verification Automation Tool

Automated workflow for processing eCW reports and uploading to HealthX portal.

---

## 📋 Overview

This tool automates the manual eCW → HealthX Excel processing workflow that was previously done using Excel macros. It:

1. **Takes** raw eCW report Excel files and template files
2. **Processes** the data using filtering rules from the original VBA macro
3. **Generates** cleaned Excel files ready for HealthX import
4. **Uploads** the processed files to HealthX portal automatically

**Before**: Manual Excel macro execution → Manual file upload  
**After**: Click button → Automatic processing → Automatic upload

---

## 🏗️ **Architecture: How Files Are Connected**

### **File Relationships Diagram**

```
┌─────────────────────────────────────────────────────────────┐
│                        main.py                              │
│  (GUI - User Interface & Workflow Orchestrator)            │
│                                                             │
│  1. User selects files → 2. Calls → 3. Calls → 4. Shows    │
│     via GUI              process_data.py   upload_hx.py    result
└────────────┬─────────────────────┬──────────┬──────────────┘
             │                     │          │
             │                     │          │
             ▼                     ▼          ▼
      ┌─────────────┐      ┌──────────────┐  ┌─────────────┐
      │ config.json │      │  config.json │  │ config.json │
      │ (paths)     │      │  (paths)     │  │ (credentials│
      └─────────────┘      └──────────────┘  │  & URL)     │
                                              └─────────────┘
             │                     │          │
             │                     │          │
             ▼                     ▼          ▼
      ┌─────────────┐      ┌──────────────┐  ┌─────────────┐
      │  inputs/    │      │  outputs/   │  │   logs/     │
      │  folder     │      │  folder     │  │   folder    │
      └─────────────┘      └──────────────┘  └─────────────┘
```

### **Data Flow**

```
User Uploads Files
        │
        ▼
   main.py (saves to inputs/)
        │
        ▼
   process_data.py
   ├── Reads eCW report (from inputs/)
   ├── Reads Help sheet (from template file)
   ├── Applies filters (3 main filters + optional)
   ├── Sorts and assigns agents
   └── Saves to outputs/
        │
        ▼
   upload_hx.py
   ├── Reads credentials from config.json
   ├── Logs into HealthX portal
   ├── Uploads file from outputs/
   └── Confirms success
        │
        ▼
   All modules log to logs/
```

---

## 📁 **Project Structure & File Details**

```
AudentesAutomation/
│
├── main.py                    # GUI application (entry point)
├── process_data.py            # Data processing engine
├── upload_hx.py               # HealthX upload automation
├── config.json                # Configuration & credentials
├── requirements.txt           # Python dependencies
├── build_exe.bat              # Windows build script
├── build_exe.sh               # Linux/Mac build script
│
├── inputs/                    # Uploaded files (auto-created)
├── outputs/                   # Processed Excel files (auto-created)
└── logs/                      # Log files (auto-created)
```

---

## 🔧 **Component Details: How Each File Works**

### **1. `main.py` - GUI Application & Orchestrator**

**Purpose**: User interface and workflow coordinator

**What it does**:
- Creates a Tkinter GUI window with file selection buttons
- Manages the complete workflow sequence
- Displays status messages to users
- Handles errors and shows success/failure popups

**Key Functions**:
```python
build_gui()              # Creates and displays the GUI window
on_run_click()           # Triggered when "Run Process" button is clicked
run_process_async()      # Runs the workflow in a background thread
```

**Workflow Steps**:
1. **File Upload**: When user clicks "Run Process"
   - Validates that both files are selected
   - Copies files to `inputs/` folder with timestamps
   - Creates a log file for this run

2. **Processing**: Calls `process_data.py`
   - Passes file paths
   - Waits for processing to complete
   - Gets the output file path and record count

3. **Upload**: Calls `upload_hx.py`
   - Passes the output file path
   - Waits for upload to complete
   - Gets success/failure status

4. **Feedback**: Shows results to user
   - Success message with record count
   - Error message if something fails
   - Updates status label throughout

**Dependencies**:
- Reads `config.json` for folder paths
- Imports `process_data.py` for data processing
- Imports `upload_hx.py` for portal upload
- Uses `tkinter` for GUI (built-in Python library)

---

### **2. `process_data.py` - Data Processing Engine**

**Purpose**: Replaces Excel macro logic with Python code

**What it does**:
- Loads and processes eCW report Excel files
- Applies filtering rules based on Help sheet lookups
- Transforms and cleans data
- Assigns allocation priorities and agents
- Generates cleaned Excel output file

**Key Functions**:
```python
generate_healthx_import()        # Main processing function
_assign_priority_and_agents()    # Assigns priorities and agents
_detect_insurance_columns()       # Finds insurance columns
_load_agents_from_config()        # Gets agent list from config
```

**Processing Steps**:

**Step 1: Load Data**
- Reads eCW report Excel file (typically has thousands of rows)
- Reads Help sheet from template file (contains lookup tables)

**Step 2: Apply Filters** (Based on VBA Macro Logic)

**Filter 1: Appointment State Validation**
- Checks if each row's " Appointment State" exists in Help sheet
- Removes rows where Appointment State is not valid
- **Logic**: Only process appointments from valid states

**Filter 2: Workable Status Check**
- Looks up each row's "Visit Type" in Help sheet
- Checks if Visit Type is marked as "Workable" (Y/N)
- Removes rows where Workable Status = "N"
- **Logic**: Only process visit types that are workable

**Filter 3: Insurance Code Exclusion**
- Excludes rows where Primary Insurance Name matches:
  - L105, L107, L109C, L109Q, L109W
- **Logic**: These insurance codes are non-billable or excluded

**Step 3: Optional Filters**
- Visit Status filter (PEN/PR) - if needed
- Empty row removal

**Step 4: Sorting**
- Sorts by: Appointment Provider Name, Appointment Date
- Ensures consistent ordering

**Step 5: Allocation Assignment**
- Groups visits by type: New Patient (NP) vs Follow-Up (FU)
- Assigns priority numbers: NP-001, NP-002, FU-001, etc.
- Distributes records across 8 agents
- Tries to keep same provider assigned to same agent

**Step 6: Save Output**
- Saves to `outputs/HealthX_Import_YYYYMMDD_HHMMSS.xlsx`
- All original columns preserved + new columns (Allocation Priority, Assigned Agent)

**Dependencies**:
- Reads `config.json` for folder paths and agent list
- Uses `pandas` for data processing
- Uses `openpyxl` for Excel file operations

---

### **3. `upload_hx.py` - HealthX Portal Automation**

**Purpose**: Automates file upload to HealthX portal using browser automation

**What it does**:
- Controls Chrome browser using Selenium
- Automatically logs into HealthX portal
- Navigates to upload page
- Uploads the processed Excel file
- Confirms successful upload

**Key Functions**:
```python
upload_to_healthx()      # Main upload function
_resolve_secret()        # Handles environment variable credentials
_load_config()           # Gets configuration from config.json
```

**Upload Steps**:

**Step 1: Browser Launch**
- Reads `headless_mode` from config.json
- Launches Chrome browser (visible or headless)
- Configures browser options

**Step 2: Login**
- Navigates to HealthX login URL (from config.json)
- Finds username and password input fields
- Enters credentials (from config.json or environment variables)
- Clicks login button
- Waits for login to complete

**Step 3: Navigate to Upload**
- Finds and clicks "Import" or "Upload" link
- Waits for upload page to load

**Step 4: Upload File**
- Finds file upload input element
- Sends the absolute path of the output file
- Finds and clicks submit/upload button
- Waits for upload to process

**Step 5: Confirm Success**
- Looks for success indicators on the page
- Returns success/failure status
- Logs the result

**Retry Logic**:
- If upload fails, retries once after 3 seconds
- Logs errors for debugging

**Dependencies**:
- Reads `config.json` for URL and credentials
- Uses `selenium` for browser automation
- Requires Chrome browser installed

**Note**: The element selectors (how it finds buttons/fields) may need to be updated if HealthX portal changes its design.

---

### **4. `config.json` - Configuration File**

**Purpose**: Stores all settings and credentials in one place

**Structure**:
```json
{
  "healthx_url": "https://hx.jindalx.com/",
  "username": "ENTER_USERNAME_HERE",
  "password": "ENTER_PASSWORD_HERE",
  "input_folder": "inputs",
  "output_folder": "outputs",
  "log_folder": "logs",
  "headless_mode": false,
  "agents": ["Agent_1", "Agent_2", ..., "Agent_8"]
}
```

**How Each Field is Used**:

- **`healthx_url`**: Used by `upload_hx.py` to navigate to HealthX portal
- **`username`**: Used by `upload_hx.py` for login
- **`password`**: Used by `upload_hx.py` for login
- **`input_folder`**: Used by `main.py` and `process_data.py` to save uploaded files
- **`output_folder`**: Used by `process_data.py` to save processed files
- **`log_folder`**: Used by all modules to save log files
- **`headless_mode`**: Used by `upload_hx.py` to run Chrome in background (true) or visible (false)
- **`agents`**: Used by `process_data.py` to assign records to agents

**Environment Variable Support**:
- If `username` = `"ENV_USERNAME"`, reads from system environment variable `USERNAME`
- If `password` = `"ENV_PASSWORD"`, reads from system environment variable `PASSWORD`
- Useful for keeping credentials out of the config file

---

### **5. `requirements.txt` - Python Dependencies**

Lists all Python packages needed:
- `pandas` - Data processing
- `openpyxl` - Excel file reading/writing
- `selenium` - Browser automation
- `python-dateutil` - Date utilities
- `pyinstaller` - Building executable

**Installation**: `pip install -r requirements.txt`

---

### **6. Build Scripts (`build_exe.bat` / `build_exe.sh`)**

**Purpose**: Package Python code into standalone executable

**What they do**:
- Use PyInstaller to create a single .exe file
- Include `config.json` with the executable
- Bundle all Python dependencies
- Creates `dist/AudentesAutomationTool.exe`

**Result**: End users don't need Python installed - just double-click the .exe

---

## 🌍 **Real-Life Usage Scenario**

### **Daily Workflow**

**Scenario**: A healthcare administrator needs to process daily eCW reports and upload them to HealthX.

**Before This Tool**:
1. Manually open Excel file with macro
2. Run Excel macro (wait 2-3 minutes)
3. Manually check results
4. Open HealthX portal in browser
5. Log in manually
6. Navigate to upload page
7. Select and upload file
8. Wait for confirmation
9. Repeat for each report

**Time**: ~10-15 minutes per report

**With This Tool**:
1. Double-click `AudentesAutomationTool.exe`
2. Click "Select eCW Report" → Choose file
3. Click "Select Template/Macro File" → Choose template
4. Click "Run Process"
5. Wait 1-2 minutes (automated)
6. See success message

**Time**: ~2-3 minutes per report (80% time savings)

**What Happens Behind the Scenes**:
1. Files saved to `inputs/` folder (with timestamps for audit trail)
2. Data processing happens automatically:
   - 3,478 raw rows → filtered to ~600 workable rows
   - Filters applied based on Help sheet rules
   - Agents assigned automatically
3. Cleaned file saved to `outputs/` folder
4. Browser opens automatically, logs in, uploads file
5. Browser closes automatically
6. All steps logged to `logs/` folder

**Multiple Runs**:
- Each run creates separate timestamped files
- Easy to track what was processed and when
- Logs provide full audit trail

---

## 🚀 **Getting Started**

### **Prerequisites**

- Python 3.8+ (if running as Python script)
- Chrome browser (required for HealthX upload)
- ChromeDriver (automatically managed by Selenium 4.x)

### **Installation**

1. **Download or clone the project**

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure `config.json`**:
   ```json
   {
     "healthx_url": "https://your-portal-url.com/login",
     "username": "your_username",
     "password": "your_password",
     "headless_mode": false
   }
   ```

### **Running the Application**

**Option 1: As Python Script**
```bash
python main.py
```

**Option 2: As Executable** (after building)
```bash
build_exe.bat          # Creates .exe
dist/AudentesAutomationTool.exe    # Run the executable
```

---

## 📊 **Data Processing Logic Explained**

### **Filter Rules (Based on VBA Macro)**

The tool applies these filters in sequence:

1. **Appointment State Filter**
   - **Rule**: Keep only rows where Appointment State exists in Help sheet
   - **Why**: Only process appointments from valid states
   - **Impact**: Removes invalid/incomplete state data

2. **Workable Status Filter** ⭐ **MOST IMPORTANT**
   - **Rule**: Keep only rows where Visit Type maps to "Workable" = "Y"
   - **Why**: Only certain visit types are billable/workable
   - **Impact**: Removes procedures, non-billable visits (typically removes ~80% of rows)

3. **Insurance Code Exclusion**
   - **Rule**: Exclude specific insurance codes: L105, L107, L109C, L109Q, L109W
   - **Why**: These codes are non-billable or excluded per business rules
   - **Impact**: Removes rows with excluded insurance types

4. **Visit Status Filter** (Optional)
   - **Rule**: Keep only "PEN" (Pending) or "PR" (Pending Referral)
   - **Why**: Only process pending appointments
   - **Note**: This filter is not in the original macro, but may be needed

### **Transformation Steps**

After filtering:
1. **Sorting**: By Provider Name, then Appointment Date
2. **Priority Assignment**: NP-001, NP-002, FU-001, etc. (based on Visit Type)
3. **Agent Distribution**: Round-robin across 8 agents (keeping same provider together when possible)

---

## 📝 **Logging System**

Each run creates a timestamped log file: `logs/run_YYYYMMDD_HHMMSS.txt`

**Log Format**:
```
[14:35:30] ============================================================
[14:35:30] Audentes Verification Automation Tool - Run Started
[14:35:30] ============================================================
[14:35:31] eCW file loaded: eCW_20251101_143022.xlsx
[14:35:32] Template file loaded: Template_20251101_143022.xlsx
[14:35:33] Loading eCW report...
[14:35:34] Loaded 3478 rows from eCW report
[14:35:35] Loading Help sheet for lookup rules...
[14:35:36] Loaded Help sheet with 193 rows
[14:35:37] Filter 1 (Appointment State in Help sheet): 3478 -> 3200 rows
[14:35:38] Filter 2 (Workable status = Y): 3200 -> 605 rows
[14:35:39] Filter 3 (Excluded insurance codes): 605 -> 600 rows
[14:35:40] Sorted by: Appointment Provider Name, Appointment Date
[14:35:41] Assigning allocation priority and agents...
[14:35:42] Saving output file: HealthX_Import_20251101_143542.xlsx
[14:35:43] Completed processing. Input rows=3478, output rows=600
[14:36:10] Starting HealthX upload process...
[14:36:15] Launching Chrome browser...
[14:36:20] Navigating to HealthX login page...
[14:36:25] Entering credentials...
[14:36:30] Uploading file...
[14:36:35] HealthX upload completed successfully
```

**Use Cases**:
- Debugging errors
- Audit trail
- Performance tracking
- Troubleshooting upload issues

---

## 🔒 **Security Notes**

- **Credentials**: Use environment variables for passwords (set in config.json)
- **Config File**: Keep `config.json` secure, don't share with real credentials
- **Logs**: May contain sensitive data - secure appropriately

---

## 🔒 Credentials Options

You can provide HealthX credentials in any of these ways (preferred → less preferred):

- Environment variables via .env (recommended)
- Environment variables via launcher script
- Plaintext in config.json (not recommended)

Note: The app supports either `user_id` (preferred) or `username`. If both are present, `user_id` is used.

### Option A: .env file (recommended)
Create a file named `.env` in the same folder as the app (or .exe):

```
ENV_HEALTHX_USER_ID=your_id
ENV_HEALTHX_PASSWORD=your_password
```

Keep `config.json` as:

```
{
  "user_id": "ENV_HEALTHX_USER_ID",
  "password": "ENV_HEALTHX_PASSWORD",
  ...
}
```

The app loads `.env` automatically (via python-dotenv) and reads those values.

### Option B: Launcher script (Windows)
Create `run.bat` in the same folder:

```
@echo off
set ENV_HEALTHX_USER_ID=your_id
set ENV_HEALTHX_PASSWORD=your_password
AudentesAutomationTool.exe
```

### Option C: Put values in config.json (not recommended)
Set `user_id` and `password` directly in `config.json`. This is less secure.

---

## 📦 **Building Executable**

### **Windows**
```bash
build_exe.bat
```
Creates: `dist/AudentesAutomationTool.exe`

### **Linux/Mac**
```bash
chmod +x build_exe.sh
./build_exe.sh
```

### **Deployment Package**
Include:
- `AudentesAutomationTool.exe`
- `config.json`
- Empty folders: `inputs/`, `outputs/`, `logs/`

---

## 📄 **Support**

For issues:
1. Check `logs/` folder for error details
2. Review this README
3. Contact development team with:
   - Log file from failed run
   - Screenshot of error (if GUI)
   - Description of what happened

---

**Version**: 1.0  
**Last Updated**: Based on VBA macro extraction and requirements analysis
