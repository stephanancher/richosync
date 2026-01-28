# RichoSync.ps1
# Generates a CSV file for Richo Address Book import from Active Directory users.

$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$IniFile = Join-Path $ScriptDir "filnavn.ini"
$LogFile = Join-Path $ScriptDir "RichoSync-data.log"
# $CsvFile is determined dynamically based on INI content

# Helper function for logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogContent = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogContent
    $Color = if ($Level -eq "ERROR") { "Red" } else { "Green" }
    Write-Host $LogContent -ForegroundColor $Color
}

# Helper to calculate Title grouping (1-10)
function Get-TitleGroup {
    param ([string]$Name)
    
    if ([string]::IsNullOrWhiteSpace($Name)) { return "1" }
    
    $FirstChar = $Name.Substring(0, 1).ToUpper()
    
    switch -Regex ($FirstChar) {
        "[A-B]" { return "1" }
        "[C-D]" { return "2" }
        "[E-F]" { return "3" }
        "[G-H]" { return "4" }
        "[I-K]" { return "5" }
        "[L-N]" { return "6" }
        "[O-Q]" { return "7" }
        "[R-T]" { return "8" }
        "[U-W]" { return "9" }
        "[X-Z]" { return "10" }
        Default { return "1" }
    }
}

# Helper to truncate string to max length (accounting for brackets)
function Get-TruncatedString {
    param (
        [string]$Value,
        [int]$MaxLength
    )
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    
    # We need to fit [Value] into MaxLength
    # So Value length must be <= MaxLength - 2
    $MaxInnerLength = $MaxLength - 2
    
    if ($MaxInnerLength -le 0) { return "" }
    
    if ($Value.Length -le $MaxInnerLength) {
        return $Value
    }
    else {
        return $Value.Substring(0, $MaxInnerLength)
    }
}

# Helper to sanitize string (remove newlines and trim)
function Sanitize-String {
    param ([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    return $Value -replace "[\r\n]", "" | ForEach-Object { $_.Trim() }
}

try {
    Write-Log "Starting RichoSync script..."

    # Check for INI file
    if (-not (Test-Path $IniFile)) {
        throw "Configuration file not found: $IniFile"
    }

    # Import AD Module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "ActiveDirectory module is not installed."
    }
    Import-Module ActiveDirectory

    # Read OUs
    $OUs = @(Get-Content $IniFile | Where-Object { $_ -and -not $_.StartsWith(";") })
    
    if ($OUs.Count -eq 0) {
        throw "No OUs found in $IniFile"
    }

    # Determine Filename Prefix from the first OU
    # Example: OU=SOE,OU=Brugere... -> Prefix: "SOE"
    $FirstOU = $OUs[0]
    $Prefix = ""
    
    if ($FirstOU -match "^OU=([^,]+)") {
        $Prefix = $Matches[1]
    }
    else {
        # Fallback if pattern doesn't match standard OU=Name format
        # Just use the beginning up to the first comma as a fallback sanitization
        $Prefix = $FirstOU.Split(",")[0] -replace "[^a-zA-Z0-9]", ""
    }

    Write-Log "Determined filename prefix: '$Prefix' from '$FirstOU'"
    
    $CsvFileName = "${Prefix}RichoUsers.csv"
    $CsvFile = Join-Path $ScriptDir $CsvFileName
    
    $UsersToRxport = @()
    $Index = 1 # Start index at 1 for the hardcoded entry

    # ---------------------------------------------------------
    # Hardcoded First Entry: Skan til Novax
    # ---------------------------------------------------------
    Write-Log "Adding hardcoded entry: Skan til Novax"
    $FirstRow = [PSCustomObject]@{
        "Index in ACLs and Groups"                    = "[$Index]"
        "Name"                                        = "[Skan til Novax]"
        "Set General Settings"                        = "[1]"
        "Set Registration No."                        = "[0]"
        "Registration No."                            = "[$Index]"
        "Entry Type"                                  = "[U]"
        "Phonetic Name"                               = "[]"
        "Display Name"                                = "[Skan til Novax]"
        "Display Priority"                            = "[5]"
        "Set Title Settings"                          = "[1]"
        "Title 1"                                     = "[8]"
        "Title 2"                                     = "[0]"
        "Title 3"                                     = "[0]"
        "Title Freq."                                 = "[1]"
        "Set User Code Settings"                      = "[0]"
        "User Code"                                   = "[]"
        "Set Auth. Info Settings"                     = "[1]"
        "Device Login User Name"                      = "[]"
        "Device Login Password"                       = "[]"
        "Device Login Password Encoding"              = "[omitted]"
        "SMTP Authentication"                         = "[0]"
        "SMTP Authentication Login User Name"         = "[]"
        "SMTP Authentication Login Password"          = "[]"
        "SMTP Authentication Password Encoding"       = "[omitted]"
        "Folder Authentication"                       = "[2]"
        "Folder Authentication Login User Name"       = "[gf\skannovax]"
        "Folder Authentication Login Password"        = "[]"
        "Folder Authentication Password Encoding"     = "[omitted]"
        "LDAP Authentication"                         = "[0]"
        "LDAP Authentication Login User Name"         = "[]"
        "LDAP Authentication Login Password"          = "[]"
        "LDAP Authentication Password Encoding"       = "[omitted]"
        "Set Access Control Settings"                 = "[0]"
        "Can Use B/W Copy"                            = "[0]"
        "Can Use Single Color Copy"                   = "[0]"
        "Can Use Two Color Copy"                      = "[0]"
        "Can Use Full Color Copy"                     = "[0]"
        "Can Use Auto Color Copy"                     = "[0]"
        "Can Use B/W Print"                           = "[0]"
        "Can Use Color Print"                         = "[0]"
        "Can Use Scanner"                             = "[]"
        "Can Use Fax"                                 = "[]"
        "Can Use Document Server"                     = "[]"
        "Maximum of Print Usage Limit"                = "[-1]"
        "Set Email/Fax Settings"                      = "[0]"
        "Fax Destination"                             = "[]"
        "Fax Line Type"                               = "[g3]"
        "International Fax Transmission Mode"         = "[]"
        "E-mail Address"                              = "[]"
        "Ifax Address"                                = "[]"
        "Ifax Enable"                                 = "[0]"
        "Direct SMTP"                                 = "[]"
        "Ifax Direct SMTP"                            = "[]"
        "Fax Header"                                  = "[1]"
        "Label Insertion 1st Line (Selection)"        = "[]"
        "Label Insertion 2nd Line (String)"           = "[]"
        "Label Insertion 3rd Line (Standard Message)" = "[0]"
        "Set Folder Settings"                         = "[1]"
        "Folder Protocol"                             = "[0]"
        "Folder Port No."                             = "[21]"
        "Folder Server Name"                          = "[]"
        "Folder Path"                                 = "[\\10.6.10.135\Novax\Skan]"
        "Folder Japanese Character Encoding"          = "[us-ascii]"
        "Set Protection Settings"                     = "[1]"
        "Is Setting Destination Protection"           = "[1]"
        "Is Protecting Destination Folder"            = "[0]"
        "Is Setting Sender Protection"                = "[0]"
        "Is Protecting Sender"                        = "[0]"
        "Sender Protection Password"                  = "[]"
        "Sender Protection Password Encoding"         = "[omitted]"
        "Access Privilege to User"                    = "[]"
        "Access Privilege to Protected File"          = "[]"
        "Set Group List Settings"                     = "[1]"
        "Groups"                                      = "[]"
        "Set Counter Reset Settings"                  = "[0]"
        "Enable Plot Counter Reset"                   = "[]"
        "Enable Fax Counter Reset"                    = "[]"
        "Enable Scanner Counter Reset"                = "[]"
        "Enable User Volume Counter Reset"            = "[]"
    }
    $UsersToRxport += $FirstRow

    foreach ($OU in $OUs) {
        Write-Log "Processing OU: $OU"
        try {
            $ADUsers = Get-ADUser -Filter { Enabled -eq $true -and EmailAddress -like "*" } -SearchBase $OU -Properties DisplayName, EmailAddress, sn, GivenName
            
            foreach ($User in $ADUsers) {
                $Index++
                
                # Sanitize inputs to prevent multi-line issues
                $RawDisplayName = Sanitize-String -Value $User.DisplayName
                $RawName = Sanitize-String -Value $User.Name
                
                $NameSource = if (-not [string]::IsNullOrWhiteSpace($RawDisplayName)) { $RawDisplayName } else { $RawName }
                
                # Truncate Name to 20 bytes (inclusive of brackets)
                $NameInner = Get-TruncatedString -Value $NameSource -MaxLength 20
                
                # Truncate Display Name to 16 bytes (inclusive of brackets)
                $DisplayNameInner = Get-TruncatedString -Value $NameSource -MaxLength 16
                
                $TitleIndex = Get-TitleGroup -Name $NameSource
                
                # Sanity check
                if (-not $TitleIndex -or $TitleIndex -eq "0") {
                    $TitleIndex = "1"
                }
                
                # Construct output object based on SDD headers and Working.csv format
                # We format values as "[Value]"
                $Row = [PSCustomObject]@{
                    "Index in ACLs and Groups"                    = "[$Index]"
                    "Name"                                        = "[$NameInner]"
                    "Set General Settings"                        = "[1]"
                    "Set Registration No."                        = "[0]"
                    "Registration No."                            = "[$Index]"
                    "Entry Type"                                  = "[U]"
                    "Phonetic Name"                               = "[]"
                    "Display Name"                                = "[$DisplayNameInner]"
                    "Display Priority"                            = "[5]"
                    "Set Title Settings"                          = "[1]"
                    "Title 1"                                     = "[$TitleIndex]" # Mapped to Grouping based on example analysis
                    "Title 2"                                     = "[0]"
                    "Title 3"                                     = "[0]"
                    "Title Freq."                                 = "[1]"
                    "Set User Code Settings"                      = "[0]"
                    "User Code"                                   = "[]"
                    "Set Auth. Info Settings"                     = "[1]"
                    "Device Login User Name"                      = "[]"
                    "Device Login Password"                       = "[]"
                    "Device Login Password Encoding"              = "[omitted]"
                    "SMTP Authentication"                         = "[0]"
                    "SMTP Authentication Login User Name"         = "[]"
                    "SMTP Authentication Login Password"          = "[]"
                    "SMTP Authentication Password Encoding"       = "[omitted]"
                    "Folder Authentication"                       = "[0]"
                    "Folder Authentication Login User Name"       = "[]"
                    "Folder Authentication Login Password"        = "[]"
                    "Folder Authentication Password Encoding"     = "[omitted]"
                    "LDAP Authentication"                         = "[0]"
                    "LDAP Authentication Login User Name"         = "[]"
                    "LDAP Authentication Login Password"          = "[]"
                    "LDAP Authentication Password Encoding"       = "[omitted]"
                    "Set Access Control Settings"                 = "[0]"
                    "Can Use B/W Copy"                            = "[0]"
                    "Can Use Single Color Copy"                   = "[0]"
                    "Can Use Two Color Copy"                      = "[0]"
                    "Can Use Full Color Copy"                     = "[0]"
                    "Can Use Auto Color Copy"                     = "[0]"
                    "Can Use B/W Print"                           = "[0]"
                    "Can Use Color Print"                         = "[0]"
                    "Can Use Scanner"                             = "[]"
                    "Can Use Fax"                                 = "[]"
                    "Can Use Document Server"                     = "[]"
                    "Maximum of Print Usage Limit"                = "[-1]"
                    "Set Email/Fax Settings"                      = "[1]"
                    "Fax Destination"                             = "[]"
                    "Fax Line Type"                               = "[g3]"
                    "International Fax Transmission Mode"         = "[0]"
                    "E-mail Address"                              = "[$($User.EmailAddress)]"
                    "Ifax Address"                                = "[]"
                    "Ifax Enable"                                 = "[0]"
                    "Direct SMTP"                                 = "[]"
                    "Ifax Direct SMTP"                            = "[]"
                    "Fax Header"                                  = "[1]"
                    "Label Insertion 1st Line (Selection)"        = "[0]"
                    "Label Insertion 2nd Line (String)"           = "[]"
                    "Label Insertion 3rd Line (Standard Message)" = "[0]"
                    "Set Folder Settings"                         = "[0]"
                    "Folder Protocol"                             = "[0]"
                    "Folder Port No."                             = "[21]"
                    "Folder Server Name"                          = "[]"
                    "Folder Path"                                 = "[]"
                    "Folder Japanese Character Encoding"          = "[us-ascii]"
                    "Set Protection Settings"                     = "[1]"
                    "Is Setting Destination Protection"           = "[1]"
                    "Is Protecting Destination Folder"            = "[]"
                    "Is Setting Sender Protection"                = "[]"
                    "Is Protecting Sender"                        = "[]"
                    "Sender Protection Password"                  = "[]"
                    "Sender Protection Password Encoding"         = "[omitted]"
                    "Access Privilege to User"                    = "[]"
                    "Access Privilege to Protected File"          = "[]"
                    "Set Group List Settings"                     = "[1]"
                    "Groups"                                      = "[]"
                    "Set Counter Reset Settings"                  = "[]"
                    "Enable Plot Counter Reset"                   = "[]"
                    "Enable Fax Counter Reset"                    = "[]"
                    "Enable Scanner Counter Reset"                = "[]"
                    "Enable User Volume Counter Reset"            = "[]"
                }
                $UsersToRxport += $Row
            }
        }
        catch {
            Write-Log "Error processing OU $OU : $_" "ERROR"
        }
    }

    # CSV Generation 
    Write-Log "Generating CSV file..."
    
    $DateFormat = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    
    # Header meta-data (Exact match to Working.csv)
    $Content = @()
    $Content += "# Format Version: 5.1.1.0"
    $Content += "# Generated at: $DateFormat"
    $Content += "# Function Name: User Data Preference"
    $Content += "# Template Name: Address Book_$DateFormat"
    $Content += '# Index in ACLs and Groups (A Number to Use as the Entry Number in CSV Files),Name (Up to 20 Characters),Set General Settings (0: Do Not Set/1: Set),Set Registration No. (0: Do Not Set/1: Set),Registration No. (Empty: Assigned Automatically/1 - 50000: Enter to Set Manually),Entry Type (U: Account/G: Group),Phonetic Name (Up to 8 Characters),Display Name (Up to 16 Characters),Display Priority (Empty: Do Not Set/1 - 10: Enter to Set Manually),Set Title Settings (0: Do Not Set/1: Set),"Title 1 (0: Do Not Register/1 - 10: ""1"" - ""10"")","Title 2 (0: Do Not Register/1: AB, 2: CD, 3: EF, 4: GH, 5: IJK, 6: LMN, 7: OPQ, 8: RST, 9: UVW, 10: XYZ)","Title 3 (0: Do Not Register/1 - 5: ""1"" - ""5"")",Title Freq. (0: Do Not Register/1: Register),Set User Code Settings (0: Do Not Set/1: Set),User Code (Up to 8 Characters),Set Auth. Info Settings (0: Do Not Set/1: Set),Login User Name (Up to 32 Characters),Device Login Password (Not Available),Device Login Password Encoding (Not Available),SMTP Authentication (0: Do Not Specify/1: Use the Login Authentication Information/2: Use Other Authentication Information),SMTP Authentication Login User Name (Up to 191 Characters),SMTP Authentication Login Password (Not Available),SMTP Authentication Password Encoding (Not Available),Folder Authentication (0: Do Not Specify/1: Use the Login Authentication Information/2: Use Other Authentication Information),Folder Authentication Login User Name (Up to 128 Characters),Folder Authentication Login Password (Not Available),Folder Authentication Password Encoding (Not Available),LDAP Authentication (0: Do Not Specify/1: Use the Login Authentication Information/2: Use Other Authentication Information),LDAP Authentication Login User Name (Up to 128 Characters),LDAP Authentication Login Password (Not Available),LDAP Authentication Password Encoding (Not Available),Set Access Control Settings (0: Do Not Set/1: Set),Can Use B/W Copy (0: Disable Black & White Copy/1: Enable Black & White Copy),Can Use Single Color Copy (0: Disable Single Color Copy/1: Enable Single Color Copy),Can Use Two Color Copy (0: Disable Two Color Copy/1: Enable Two Color Copy),Can Use Full Color Copy (0: Disable Full Color Copy/1: Enable Full Color Copy),Can Use Auto Color Copy (0: Disable Auto Color Copy/1: Enable Auto Color Copy),Can Use B/W Print (0: Disable Black & White Printing/1: Enable Black & White Printing),Can Use Color Print (0: Disable Color Printing/1: Enable Color Printing),Can Use Scanner (0: Restrict Scanner Usage/1: Do Not Restrict),Can Use Fax (0: Restrict Fax Usage/1: Do Not Restrict),Can Use Document Server (0: Restrict Document Box Usage/1: Do Not Restrict),Maximum of Print Usage Limit (Empty: Do Not Limit/0 - 999999: Enter Limit Value),Set Email/Fax Settings (0: Do Not Set/1: Set),Fax Destination (Up to 512 Characters),"Fax Line Type (g3, ext (G3 internal line), g4, g4_ext (G4 internal line), ig3, ig3_ext (I-G3 internal line), g3_auto (G3 unused line), ext_auto (G3 unused line, internal line), g3_1, g3_1_ext (G3-1 internal line), g3_2, g3_2_ext (G3-2 internal line), g3_3, g3_3_ext (G3-3 internal line), h323, sip)",International Fax Transmission Mode (0: Disable/1: Enable),E-mail Address (Up to 128 Characters),Ifax Address (Up to 128 Characters),Ifax Enable (0: E-mail and Internet Fax/1: Internet Fax Only),Direct SMTP (0: Send via SMTP Server/1: Do Not Send via SMTP Server),Ifax Direct SMTP (0: Send via SMTP Server/1: Do Not Send via SMTP Server),Fax Header (0: Do Not Set/1 - 10: Name 1 - 10),Label Insertion 1st Line (Selection) (0: Do Not Use Merge Print/1: Use Merge Print),Label Insertion 2nd Line (String) (Up to 28 Characters),Label Insertion 3rd Line (Standard Message) (0: Do Not Print/1 - 4: Print Pre-Registered Text 1 - 4),Set Folder Settings (0: Do Not Set/1: Set),Folder Protocol (0: SMB/1: FTP/2: NCP-Bindery/3: NCP-NDS),Folder Port No. (1 - 65535),Folder Server Name (Up to 128 Characters),Folder Path (Up to 256 Characters),Folder Japanese Character Encoding (us-ascii/shift_jis/euc-jp),Set Protection Settings (0: Do Not Configure/1: Configure the Settings),Is Setting Destination Protection (0: Do Not Use/1: Use the Entry as a Destination),Is Protecting Destination Folder (0: Do Not Protect/1: Protect),Is Setting Sender Protection (0: Do Not Use/1: Use the Entry as the Sender),Is Protecting Sender (0: Do Not Protect/1: Protect),Sender Protection Password (Not Available),Sender Protection Password Encoding (Not Available),"Access Privilege to User (""Index in ACLs and Groups"" Number and One of the Following Letters; R: Viewing Allowed/W: Editing Allowed/D: Editing/Deleting Allowed/X: Full Control) (e.g. 10R, 20X)","Access Privilege to Protected File (""Index in ACLs and Groups"" Number and One of the Following Letters; R: Viewing Allowed/W: Editing Allowed/D: Editing/Deleting Allowed/X: Full Control) (e.g. 10R, 20X)",Set Group List Settings (0: Do Not Configure/1: Configure the Settings),"Groups (Enter the ""Index in ACLs and Groups"" Number to Specify the Assigned Group)",Set Counter Reset Settings (0: Do Not Configure/1: Configure the Settings),"Enable Plot Counter (Copier, Printer, Fax) Reset (0: Do Not Reset/1: Reset the Counter)",Enable Fax Counter Reset (0: Do Not Reset/1: Reset the Counter),Enable Scanner Counter Reset (0: Do Not Reset/1: Reset the Counter),Enable User Volume Counter Reset (0: Do Not Reset/1: Reset the Counter),'
    $Content += "# Authentication Method (0=none or user code/1=others): 0"
    
    $Headers = "Index in ACLs and Groups", "Name", "Set General Settings", "Set Registration No.", "Registration No.", "Entry Type", "Phonetic Name", "Display Name", "Display Priority", "Set Title Settings", "Title 1", "Title 2", "Title 3", "Title Freq.", "Set User Code Settings", "User Code", "Set Auth. Info Settings", "Device Login User Name", "Device Login Password", "Device Login Password Encoding", "SMTP Authentication", "SMTP Authentication Login User Name", "SMTP Authentication Login Password", "SMTP Authentication Password Encoding", "Folder Authentication", "Folder Authentication Login User Name", "Folder Authentication Login Password", "Folder Authentication Password Encoding", "LDAP Authentication", "LDAP Authentication Login User Name", "LDAP Authentication Login Password", "LDAP Authentication Password Encoding", "Set Access Control Settings", "Can Use B/W Copy", "Can Use Single Color Copy", "Can Use Two Color Copy", "Can Use Full Color Copy", "Can Use Auto Color Copy", "Can Use B/W Print", "Can Use Color Print", "Can Use Scanner", "Can Use Fax", "Can Use Document Server", "Maximum of Print Usage Limit", "Set Email/Fax Settings", "Fax Destination", "Fax Line Type", "International Fax Transmission Mode", "E-mail Address", "Ifax Address", "Ifax Enable", "Direct SMTP", "Ifax Direct SMTP", "Fax Header", "Label Insertion 1st Line (Selection)", "Label Insertion 2nd Line (String)", "Label Insertion 3rd Line (Standard Message)", "Set Folder Settings", "Folder Protocol", "Folder Port No.", "Folder Server Name", "Folder Path", "Folder Japanese Character Encoding", "Set Protection Settings", "Is Setting Destination Protection", "Is Protecting Destination Folder", "Is Setting Sender Protection", "Is Protecting Sender", "Sender Protection Password", "Sender Protection Password Encoding", "Access Privilege to User", "Access Privilege to Protected File", "Set Group List Settings", "Groups", "Set Counter Reset Settings", "Enable Plot Counter Reset", "Enable Fax Counter Reset", "Enable Scanner Counter Reset", "Enable User Volume Counter Reset"
    
    # Header Row
    $HeaderLine = ($Headers | ForEach-Object { "`"$_`"" }) -join ","
    $Content += $HeaderLine
    
    # Data Rows
    foreach ($Row in $UsersToRxport) {
        $Line = ($Headers | ForEach-Object { "`"$($Row.$_)`"" }) -join ","
        $Content += $Line
    }
    
    $Content | Set-Content $CsvFile -Encoding UTF8
    
    Write-Log "Successfully generated $CsvFile with $($UsersToRxport.Count) users."

}
catch {
    Write-Log "Fatal Error: $_" "ERROR"
    exit 1
}
