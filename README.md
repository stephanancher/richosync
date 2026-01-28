# RichoSync

PowerShell script to generate address book CSV files for Ricoh printers by synchronizing with Active Directory.

## Purpose
This script automates the creation of a CSV file compatible with Ricoh's address book import format. It fetches users from specified Active Directory Organizational Units (OUs), formats the data (truncating names, assigning title groups), and includes specific hardcoded entries required for the system (e.g., "Skan til Novax").

## Prerequisites
- **PowerShell**
- **Active Directory Module for Windows PowerShell**
- Access to the Active Directory environment.

## Configuration
The script looks for a configuration file named `filnavn.ini` in the same directory.
This file should contain the Distinguished Names (DN) of the OUs to scan, one per line. Lines starting with `;` are ignored as comments.

**Example `filnavn.ini`:**
```ini
; List of OUs to scan
OU=Users,OU=Department,DC=example,DC=com
OU=IT,OU=Staff,DC=example,DC=com
```

## Usage
Run the script directly from PowerShell:

```powershell
.\RichoSync.ps1
```

## Output
The script generates a CSV file named based on the first OU found in the configuration (e.g., `SOERichoUsers.csv` or `SEKRichoUsers.csv`).
- **Output File**: A comma-separated values file containing user data formatted for Ricoh Address Book import.
- **Log File**: `RichoSync-data.log` contains execution details and errors.

## Key Features
- **Hardcoded Entries**: Automatically adds "Skan til Novax" as the first entry (Index 1).
- **Title Grouping**: Automatically assigns users to "Title 1" groups (1-10) based on the first letter of their name.
- **Data Sanitization**: Truncates names to meet Ricoh's character limits and removes invalid characters.
- **Dynamic Filenaming**: Determines the output filename based on the OU structure.