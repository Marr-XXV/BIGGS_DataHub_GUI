# BIGGS DataHub GUI

A comprehensive desktop application for fetching, validating, and managing BIGGS (Banking Industry Business Growth System) sales data with an intuitive graphical interface.

## Description

BIGGS DataHub GUI is a Python-based Tkinter application that enables users to:
- Fetch real-time sales data from multiple POS (Point-of-Sale) systems
- Append or rewrite data into a master CSV database
- Validate data quality with automated checks
- Search and view historical records with pagination
- Track fetch progress with live activity logs

## Features

✨ **Data Fetching**
- Multi-branch data retrieval from remote sources
- Date range filtering (YYYY-MM-DD format)
- Real-time progress monitoring with speed and ETA calculations
- Automatic duplicate detection before import

🔍 **Data Validation**
- Missing column detection
- Invalid branch code validation
- Negative value detection (invalid sales records)
- Detailed error reports with issue categorization

📊 **Record Management**
- Browse master file or dated record folders
- Multi-criteria search (branch, date range)
- Paginated view (1000 records per page)
- Previous/Next navigation with page indicators

🎨 **User Interface**
- Clean, modern design with blurred backgrounds
- Slide-out sidebar navigation
- Live fetch activity monitor tree
- Responsive layout that adapts to window resizing
- Status bar with progress bar and detailed metrics

⚙️ **Configuration**
- Editable branch list (`settings/branches.txt`)
- Current record tracking (`settings/current_record.txt`)
- Append or Rewrite mode selection
- Automatic branch suggestion dropdown

## Technologies Used

- **Python 3.8+**
- **Tkinter** - GUI framework (built-in with Python)
- **ttk** - Modern themed widgets
- **PIL/Pillow** - Image processing (logo, backgrounds, blurring)
- **Pandas** - Data validation and manipulation
- **CSV** - Data storage and exchange
- **Threading** - Asynchronous fetch operations
- **OS/datetime/time** - System operations and time tracking

## Installation / Setup

### Prerequisites
```bash
# Python 3.8 or higher
python --version

# Required packages
pip install pillow pandas
```

### Step 1: Clone or Download
```bash
# Navigate to the project directory
cd BIGGS_DataHub_GUI
```

### Step 2: Create Settings Directory
```bash
# Ensure settings folder exists
mkdir -p settings
```

### Step 3: Create Configuration Files

**settings/branches.txt** - List of approved branch codes
```
BRN001
BRN002
BRN003
```

**settings/current_record.txt** - Path to the current master file (optional)
```
record2025.csv
```

### Step 4: Prepare Image Files (Optional)
- `BIGGS_LOGO.png` - Logo for sidebar (50-100 pixels wide)
- `BIGGS_IMAGE.jpg` - Background for record viewer (landscape orientation)

### Step 5: Run the Application
```bash
python BIGGS_DATAHUB.py
```

## Usage

### Main Screens

**🏠 Home Screen**
- Click **Fetch Data** to import new records
- Click **Records** to search historical data
- Use hamburger menu (left edge) to navigate

**📥 Fetch Data Screen**
1. Enter **Start Date** and **End Date** (YYYY-MM-DD)
2. Select **Branch** from dropdown or type to filter
3. Choose **Mode**:
   - **Append**: Save as new dated file (e.g., `record_2025-12-15_to_2025-12-17.csv`)
   - **Rewrite**: Merge into `record2025.csv` (shows preview first)
4. Click **Fetch** to start download
   - Monitor progress in Live Activity Log on the right
   - Click **Cancel** to stop at any time
5. System automatically validates data and detects duplicates

**📋 Records Screen**
1. Select **Source**: Master File or Records Folder
2. Apply **Filters**:
   - Branch (optional)
   - Start Date (optional)
   - End Date (optional)
3. Click **Search** to filter results
4. Use **Prev/Next** buttons to navigate pages
5. Click **Clear** to reset and show all records

**✅ Validation**
- Access via Home Screen → Validate Records
- Checks for:
  - Missing required columns (DATE, BRANCH, POS, QUANTITY, AMOUNT)
  - Invalid branch codes
  - Negative QUANTITY or AMOUNT values
- Generates detailed report (CSV format) with all issues found

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Show Sidebar | Move mouse to left edge |
| Hide Sidebar | Move mouse away |
| Resign/Cancel | Click button or press Escape |

### Configuration

**Change Master File Location**
- Edit `settings/current_record.txt`
- Enter relative or absolute path to CSV file
- Falls back to `record2025.csv` if not found

**Update Approved Branches**
- Edit `settings/branches.txt`
- One branch code per line
- Used for validation and dropdown suggestions

## Output Files

| File | Purpose |
|------|---------|
| `record2025.csv` | Master data file (append/rewrite target) |
| `records/` folder | Dated CSV files from append-mode fetches |
| `validation_errors_YYYY-MM-DD.csv` | Validation issue report |
| `latest/` folder | Staging directory for downloads in progress |

## Troubleshooting

**"Pandas library is required for validation"**
- Install: `pip install pandas`

**"BIGGS_LOGO.png not found"**
- Optional: Provide the file or application will show text instead

**"Branch not found"**
- Update `settings/branches.txt` with the correct branch codes

**Data Not Loading in Records**
- Verify CSV file exists and path in `settings/current_record.txt` is correct
- Check file has required columns: DATE, BRANCH, POS, QUANTITY, AMOUNT

**Fetch Operation Hangs**
- Click **Cancel** button
- Check internet connection and remote server status
- Review log messages in Activity Monitor

## File Structure

```
BIGGS_DataHub_GUI/
├── BIGGS_DATAHUB.py          # Main GUI application
├── fetcher.py                # Remote data fetch module
├── combiner.py               # CSV merge/combine logic
├── pandasbiggs.py            # Data processing utilities
├── missing_generate.py       # Missing data handler
├── manual_fetch.py           # Manual fetch utility
├── README.md                 # This file
├── settings/
│   ├── branches.txt          # Approved branch codes
│   ├── current_record.txt    # Master file path
│   └── newBranches.txt       # (Auto-generated)
├── records/                  # Historical dated records
├── latest/                   # Staging directory
└── BIGGS_LOGO.png           # (Optional) Logo image
BIGGS_IMAGE.jpg              # (Optional) Background image
```

## Tips & Best Practices

📌 **Before First Use**
- Set up `settings/branches.txt` with your branch codes
- Copy or verify the master CSV file location
- Test with a small date range first

📌 **Data Entry**
- Always use YYYY-MM-DD format for dates
- Branch codes must match `settings/branches.txt` exactly
- Select "Append" mode for regular updates; "Rewrite" only for corrections

📌 **Validation**
- Run validation after every fetch to catch issues early
- Review error reports and fix data before re-importing

📌 **Performance**
- Fetching large date ranges may take time (monitor activity log for progress)
- Record viewer loads 1000 records per page (use filters to narrow results)
- Application remains responsive during fetch (uses background threads)

📌 **Backup**
- Always backup `record2025.csv` before using Rewrite mode
- Archive old records regularly to keep file size manageable

## Author

**Developer:** Seriemar R. Arroyo  
**Collaborator:** Don John Daryl P. Curativo

## License

Proprietary - All rights reserved  
© 2025 BIGGS System. For authorized use only.

---

**Last Updated:** April 2025  
**Version:** 1.0.0

For support or questions, contact the development team.
