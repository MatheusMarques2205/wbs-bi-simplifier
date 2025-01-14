# WBS BI Simplifier

## ⚠️ Important Note
This tool is **Windows-only** as it requires Microsoft Access and its ODBC drivers. It will not work on Mac OS or Linux systems.

## About
Tool to simplify WBS data from MS Project for Power BI visualization, focused on wind energy projects. Transforms thousands of tasks into 19 main milestones.

## System Requirements
- **Operating System**: Windows 10 or later (mandatory)
- **Database**: Microsoft Access 2010 or later
- **Python**: Version 3.6 or later
- **Power BI**: Desktop version (latest recommended)
- **RAM**: Minimum 4GB
- **Storage**: 500MB free disk space
- **Microsoft Office**: 32 or 64-bit version (for Access Database Engine)

## Detailed Installation Guide

### 1. Python Installation
1. Download Python:
   - Visit [Python.org](https://www.python.org/downloads/)
   - Download Python 3.x (latest stable version)
   - **Important**: Check "Add Python to PATH" during installation

2. Verify Installation:
   - Open Command Prompt (cmd)
   - Type: `python --version`
   - Should display Python version

### 2. Required Extensions and Packages

#### Install pip (if not installed):
```bash
python -m ensurepip --default-pip
```

#### Install required packages:
```bash
pip install -r requirements.txt
```

#### Verify installations:
```bash
pip list
```

### 3. Microsoft Access Driver (Windows Only)
1. Verify ODBC Driver:
   - Open Windows Search
   - Type "ODBC Data Sources"
   - Check if "Microsoft Access Driver (*.mdb, *.accdb)" is listed
2. If driver is missing:
   - Download "Microsoft Access Database Engine"
   - Install the version matching your Office installation (32 or 64-bit)
   - Download link: [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

## How to Use

### Step by Step

1. Configure Python file:
   - Open "WBS (Work Breakdown Structure).py"
   - Update database path:
   ```python
   DATABASE_PATH = "your_database_path.mdb"
   ```

2. Run the program:
   - Open terminal
   - Run: `python "WBS (Work Breakdown Structure).py"`

3. In Power BI:
   - Import both tables using Power Query
   - Create relationship between tables using TASK_ID
   - Create your visualizations

## Main Milestones
The program categorizes tasks into:
1. NTP
2. Anchor Cages EX Works
3. Anchor Cages Delivery at POD
4. Anchor Cages Delivery at Site
5. WTG Ex Works Local
6. Foundation Ready
7. WTG Delivery at Site
8. Main Crane Mobilization
9. Pre Erection
10. Full Erection
11. Main Crane Demobilization
12. Electrical and Mechanical Completion
13. SCADA Installation
14. Pre Commissioning
15. Substation Energized
16. Commissioning
17. Run Test
18. COD

## Troubleshooting

### Common Issues and Solutions

1. Python Path Error:
   - Verify Python is in System PATH
   - Restart command prompt after installation

2. Package Installation Errors:
   ```bash
   # If pip install fails, try:
   python -m pip install --upgrade pip
   python -m pip install package_name
   ```

3. ODBC Driver Issues:
   - Ensure correct driver version (32/64-bit)
   - Reinstall Microsoft Access Database Engine if needed
   - Make sure you're using Windows OS

4. Database Connection Error:
   - Check database path
   - Verify file permissions
   - Ensure database is not locked by another program

## Support
For questions or issues:
1. Check troubleshooting guide above
2. Open an issue on GitHub
3. Include error messages and system details in reports

## Platform Compatibility
- ✅ Windows 10
- ✅ Windows 11
- ❌ MacOS (Not Compatible)
- ❌ Linux (Not Compatible)
