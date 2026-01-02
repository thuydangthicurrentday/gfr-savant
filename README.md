# GOFILEROOM Downloader

Automation script to download documents from GOFILEROOM for multiple clients.

## Features

- Automated client search and document download from GOFILEROOM
- Batch processing of multiple clients from Excel file
- Automatic document categorization into folders (Permanent, Tax, Accounting & Payroll, etc.)
- CSV export of document lists
- Single and multiple document download support
- Retry mechanism for failed downloads
- Email notifications for critical errors
- Detailed logging and Excel status tracking
- Automatic folder organization by client name and number

## Prerequisites

- Python 3.12
- Google Chrome browser
- Excel file with client list (see Configuration section)

## Installation

### 1. Install Python 3.12

Download and install Python 3.12 from [python.org](https://www.python.org/downloads/)

### 2. Open Project and Navigate to Terminal

- Open the project folder in your code editor
- Open terminal/command prompt in the project directory

### 3. Create Virtual Environment

```bash
python -m venv venv
```

### 4. Activate Virtual Environment

**For Windows PowerShell:**
```powershell
.\venv\Scripts\Activate
```

**Note:** If you get an execution policy error in PowerShell, run:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**For Windows Command Prompt:**
```cmd
venv\Scripts\activate
```

**For Linux/Mac:**
```bash
source venv/bin/activate
```

### 5. Install Required Packages

```bash
pip install -r requirements.txt
```

## Configuration

### 1. Create .env File

Create a `.env` file in the project root directory with the following variables:

```env
# Excel Configuration
CLIENT_LIST_FILE_NAME=download_gofileroom_data.xlsx
CLIENT_LIST_SHEET_NAME=Client List GFR
DOCUMENT_LIST_SHEET_NAME=Download Document Log
NUMBER_ITEMS_PER_PAGE=50

# Login Credentials (REQUIRED - update with your GOFILEROOM account credentials)
USERNAME=your_email@example.com
PASSWORD=your_password

# Download Directory (REQUIRED - change to your desired download path)
# Use absolute path or relative path from project root
# Example: DOWNLOAD_DIR=C:\Users\YourName\Downloads\gofileroom

# Error Handling (optional)
MAX_CONSECUTIVE_ERRORS=10
DOWNLOAD_RETRY_COUNT=3

# Email Configuration (optional - for error notifications)
ENABLE_EMAIL=True
EMAIL_HOST=smtp.gmail.com
EMAIL_PORT=587
EMAIL_HOST_USER=your_email@gmail.com
EMAIL_HOST_PASSWORD=your_app_password
EMAIL_USE_TLS=True
DEFAULT_FROM_EMAIL=your_email@gmail.com
EMAIL_recipient_list=recipient1@example.com,recipient2@example.com
MACHINE=Machine-Name
```

**Important Notes:** 
- **REQUIRED**: Replace all placeholder values with your actual credentials and configuration
- **USERNAME and PASSWORD**: You **must** update `USERNAME` and `PASSWORD` with your GOFILEROOM account credentials:
  - `USERNAME` should be your email address used to login to GOFILEROOM
  - `PASSWORD` should be your GOFILEROOM account password
  - Example: `USERNAME=yourname@company.com` and `PASSWORD=your_actual_password`
- **DOWNLOAD_DIR**: You **must** change `DOWNLOAD_DIR` to your desired download directory path. This is where all downloaded files will be stored. You can use either:
  - Absolute path: `DOWNLOAD_DIR=C:\Users\YourName\Downloads\gofileroom`
  - Relative path: `DOWNLOAD_DIR=./downloads` (relative to project root)
- Email configuration is optional. Set `ENABLE_EMAIL=False` to disable email notifications
- For Gmail email notifications, you may need to use an [App Password](https://support.google.com/accounts/answer/185833) instead of your regular password

### 2. Configure Chrome Download Location

**IMPORTANT**: You must configure Chrome's download location to match the `DOWNLOAD_DIR` value in your `.env` file.

1. Open Chrome browser
2. Click the three dots menu (⋮) in the top right corner
3. Go to **Settings** → **Downloads** (or type `chrome://settings/downloads` in the address bar)
4. Under **Location**, click **Change**
5. Set the download folder to **exactly the same path** as your `DOWNLOAD_DIR` in the `.env` file
   - If `DOWNLOAD_DIR=C:\Users\YourName\Downloads\gofileroom`, set Chrome download location to `C:\Users\YourName\Downloads\gofileroom`
   - If `DOWNLOAD_DIR=./downloads`, set Chrome download location to the full absolute path (e.g., `C:\path\to\project\downloads`)
6. **Important**: Make sure the paths match exactly (including case sensitivity on Linux/Mac)

**Note**: The script uses Selenium to monitor the download directory, so Chrome must download files to the same location specified in `DOWNLOAD_DIR`.

### 3. Prepare Excel File

1. Ensure the Excel file specified in `CLIENT_LIST_FILE_NAME` exists in the project root directory
2. The Excel file must contain two sheets:
   - **Client List GFR** (or name specified in `CLIENT_LIST_SHEET_NAME`)
   - **Download Document Log** (or name specified in `DOCUMENT_LIST_SHEET_NAME`)
3. In the **Client List GFR** sheet, set column A (Status) to **"Pending"** for clients you want to process
4. The sheet should have the following columns:
   - Status
   - Description
   - Client Name
   - Client Number
   - Total Documents
   - Number Of Files Downloaded
   - Client Folder Path
   - Client Email (optional)

## Running the Script

### 1. Configure Chrome Download Location (if not done already)

Before starting Chrome with remote debugging, ensure Chrome's download location matches your `DOWNLOAD_DIR` from the `.env` file (see Configuration section above).

### 2. Start Chrome with Remote Debugging

Before running the main script, you need to start Chrome with remote debugging enabled. Open a new terminal/command prompt and run:

**For Windows:**
```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp_chrome_data"
```

**Important Notes:** 
- **This is your main Chrome browser instance** that will be used by the automation script
- **First time setup**: When you run this command for the first time, Chrome will be initialized with a new profile stored in `C:\temp_chrome_data`. This profile will be reused for all subsequent automation runs
- **Login to GOFILEROOM**: After Chrome opens, manually login to GOFILEROOM (https://www.gofileroom.com/login) using your credentials to verify that login works correctly. This login session will be saved and reused for future automation runs
- **Setup Download Location**: In this Chrome instance, go to Settings → Downloads and set the download location to match your `DOWNLOAD_DIR` from the `.env` file (see Configuration section above)
- **Keep Chrome open**: Keep this Chrome window open while running the automation script. Do not close it
- **Reuse for future runs**: For subsequent automation runs, you can reuse this same Chrome instance - just run the command again and Chrome will open with your saved profile (including login session and settings)
- If Chrome is installed in a different location, adjust the path accordingly
- The `temp_chrome_data` directory will be created automatically on first run

### 3. Verify Excel File Setup

- Open your Excel file
- Go to the **Client List GFR** sheet
- Ensure column A (Status) is set to **"Pending"** for clients you want to process
- Save and close the Excel file (important: the file must be closed before running the script)

### 4. Run the Main Script

In your activated virtual environment terminal, run:

```bash
python main.py
```

## How It Works

1. The script reads the Excel file and finds all clients with Status = "Pending"
2. For each client:
   - Searches for the client in GOFILEROOM
   - Downloads the CSV file containing document list
   - Downloads all documents (single or multiple based on document count)
   - Updates the Excel file with progress and status
3. Status updates:
   - **Pending**: Client is ready to be processed
   - **Success**: Client processed successfully
   - **Error**: Error occurred during processing (details in Description column)
   - **Warning**: Client found but has no documents

## Troubleshooting

### Excel File Error: "Invalid XML"

If you encounter an error about invalid XML when loading the Excel file:

1. **Close Excel** if the file is currently open
2. **Open the file in Excel** and save it again (File > Save As > Excel Workbook)
3. **Remove any VBA macros** or unsupported features
4. **Check for file corruption** - try opening in Excel first

### Chrome Remote Debugging Issues

- Make sure Chrome is started with the remote debugging command before running the script
- Close all existing Chrome instances before starting with remote debugging
- If port 9222 is already in use, you can change it (and update the code accordingly)

### Login Issues

- Verify your username and password in the `.env` file
- Check if your account requires 2FA (two-factor authentication)
- Ensure the login URL is correct

### File Download Issues

- Ensure the download directory exists and is writable
- Check disk space availability
- Verify Chrome download settings allow automatic downloads

## Log Files

The script creates a log file `gofileroom_download.log` in the project root directory. Check this file for detailed execution logs and error information.

## Project Structure

Key files in the project:
- `main.py` - Main automation script
- `models.py` - Client and Document data models
- `excel_handler.py` - Excel file operations
- `file_handler.py` - File system operations (move, rename, extract, etc.)
- `email_handler.py` - Email notification functionality
- `document_mapping.py` - Document categorization logic
- `config.py` - Selenium and web configuration
- `utils.py` - Utility functions (config loading, path handling)
- `.env` - Configuration file (create this, not included in repo)
- `requirements.txt` - Python dependencies

## Notes

- The script processes clients sequentially
- If an error occurs, the script will log it and continue with the next client
- After processing, check the Excel file for updated status and document counts
- Downloaded files are organized in folders by client name and number
- Documents are automatically categorized into subfolders:
  - `_Permanent` - Permanent documents
  - `Tax` - Tax-related documents
  - `Accounting & Payroll` - Accounting and payroll documents
  - `Consulting & Special Projects` - Consulting documents
  - `Other` - Other document types

## Git Setup and Push Code

### Initial Setup (First Time Only)

1. **Initialize Git Repository** (if not already initialized):
   ```bash
   git init
   ```

2. **Add Remote Repository**:
   ```bash
   git remote add origin <your-repository-url>
   ```
   Example:
   ```bash
   git remote add origin https://github.com/yourusername/gofileroom-downloader.git
   ```

3. **Configure Git User** (if not already configured):
   ```bash
   git config --global user.name "Your Name"
   git config --global user.email "your.email@example.com"
   ```

### Push Code to Repository

1. **Check Status** - See what files have changed:
   ```bash
   git status
   ```

2. **Add Files to Staging**:
   ```bash
   # Add all changes
   git add .
   
   # Or add specific files
   git add main.py README.md .gitignore
   ```

3. **Commit Changes**:
   ```bash
   git commit -m "Your commit message describing the changes"
   ```
   Example:
   ```bash
   git commit -m "Add README and improve error handling"
   ```

4. **Push to Remote Repository**:
   ```bash
   # First time pushing to a new branch
   git push -u origin main
   
   # Or if your default branch is 'master'
   git push -u origin master
   
   # Subsequent pushes
   git push
   ```

### Common Git Commands

- **View changes**: `git diff`
- **View commit history**: `git log`
- **Create new branch**: `git checkout -b feature-name`
- **Switch branch**: `git checkout branch-name`
- **Pull latest changes**: `git pull`
- **View remote repositories**: `git remote -v`

### Important Notes

- **Never commit sensitive files**: The `.gitignore` file is configured to exclude:
  - `.env` file (contains passwords and credentials)
  - Excel files with client data
  - Download directories
  - Log files
  - Virtual environment (`venv/`)

- **Before pushing, always check**:
  ```bash
  git status
  ```
  Make sure no sensitive files (like `.env` or Excel files with data) are being committed.

- **If you accidentally added sensitive files**, remove them:
  ```bash
  git rm --cached .env
  git commit -m "Remove sensitive file"
  ```

