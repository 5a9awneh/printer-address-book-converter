<#
.SYNOPSIS
    Printer Address Book Converter v1.5

.DESCRIPTION
    Converts printer address book files between Canon, Sharp, Xerox, and Develop/Konica formats.
    Auto-detects source format and handles validation, deduplication, and backups.

.PARAMETER SourcePath
    Path to source CSV file (for non-interactive mode)

.PARAMETER TargetBrand
    Target brand format: Canon, Sharp, Xerox, or Develop (for non-interactive mode)

.PARAMETER Mode
    Conversion mode: Single, Batch, or Merge (default: Single)

.PARAMETER NoInteractive
    Run in non-interactive mode without GUI prompts

.EXAMPLE
    .\Convert-PrinterAddressBook.ps1
    Interactive mode with menu navigation

.EXAMPLE
    .\Convert-PrinterAddressBook.ps1 -SourcePath "export.csv" -TargetBrand "Canon" -NoInteractive
    Non-interactive conversion to Canon format

.NOTES
    Author: Faris Khasawneh
    Created: January 2026
    Version: 1.5
    Supports: Canon (iR-ADV, imageFORCE), Sharp MX/BP, Xerox AltaLink/VersaLink, Develop/Konica/Bizhub

.CHANGELOG
    v1.5 - CRITICAL FIX: Canon CSV format now uses UNQUOTED headers with trailing comma
           Matches native Canon export format (imageFORCE-6160 compatibility)
           Ensures universal compatibility across all Canon models
    v1.4 - CRITICAL FIX: Added all missing columns to match original export formats
           Sharp: Added 15 columns (FTP/SMB/Desktop auth fields)
           Xerox: Added 13 columns (Scan/InternetFax fields)  
           Develop: Added 33 columns (FTP/SMB/WebDAV/Fax fields)
    v1.3 - Fixed Canon output: Added missing "# DB Version: 0x010b" header + blank line
    v1.2 - Added smart SearchKey derivation for Develop format
    v1.1 - Fixed Canon read, Xerox name splitting
    v1.0 - Initial release
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
    [string]$TargetBrand,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Single', 'Batch', 'Merge')]
    [string]$Mode = 'Single',

    [Parameter(Mandatory = $false)]
    [switch]$NoInteractive
)

#region Configuration

$Script:BrandConfig = @{
    'Canon'   = @{
        NameField        = 'cn'
        EmailField       = 'mailaddress'
        Encoding         = 'UTF8'
        Delimiter        = ','
        HasComments      = $true
        SignatureColumns = @('objectclass', 'cn', 'mailaddress')
        # Output field mappings
        OutputFields     = @{
            DisplayName = 'cn'
            Email       = 'mailaddress'
        }
    }
    'Sharp'   = @{
        NameField        = 'name'
        EmailField       = 'mail-address'
        Encoding         = 'UTF8'
        Delimiter        = ','
        HasComments      = $false
        SignatureColumns = @('address', 'name', 'mail-address', 'ftp-host')
        # Output field mappings
        OutputFields     = @{
            DisplayName = 'name'
            Email       = 'mail-address'
        }
    }
    'Xerox'   = @{
        NameField        = 'DisplayName'
        NameFieldAlt     = @('FirstName', 'LastName')
        EmailField       = 'E-mailAddress'
        Encoding         = 'UTF8'
        Delimiter        = ','
        HasComments      = $false
        SignatureColumns = @('XrxAddressBookId', 'DisplayName', 'E-mailAddress')
        # Output field mappings
        OutputFields     = @{
            DisplayName = 'DisplayName'
            FirstName   = 'FirstName'
            LastName    = 'LastName'
            Email       = 'E-mailAddress'
        }
    }
    'Develop' = @{
        NameField        = 'Name'
        EmailField       = 'MailAddress'
        Encoding         = 'Unicode'
        Delimiter        = "`t"
        HasComments      = $false
        SkipRows         = 2
        SignatureColumns = @('AbbrNo', 'Name', 'MailAddress')
        # Output field mappings
        OutputFields     = @{
            DisplayName = 'Name'
            Email       = 'MailAddress'
        }
    }
}

$Script:Stats = @{
    Converted     = 0
    Skipped       = 0
    Duplicates    = 0
    InvalidEmails = @()
}

$Script:LogFile = "converter-$(Get-Date -Format 'yyyy-MM-dd').log"

#endregion

#region Utility Functions

function Write-Log {
    <#
    .SYNOPSIS
        Enhanced logging with console output support
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [string]$Function = '',
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $functionPart = if ($Function) { " [$Function]" } else { '' }
    $logEntry = "$timestamp $Level$functionPart $Message"

    # Add error details if provided
    if ($ErrorRecord) {
        $logEntry += "`n  Exception: $($ErrorRecord.Exception.Message)"
        $logEntry += "`n  Location: $($ErrorRecord.InvocationInfo.ScriptName):$($ErrorRecord.InvocationInfo.ScriptLineNumber)"
    }

    try {
        Add-Content -Path $Script:LogFile -Value $logEntry -ErrorAction Stop
    }
    catch {
        # Silently ignore log write failures to prevent script interruption
        # Log file might be locked, missing permissions, or disk full
    }

    # Also write to console if -Verbose or level is WARN/ERROR
    if ($VerbosePreference -eq 'Continue' -or $Level -in @('WARN', 'ERROR')) {
        $color = switch ($Level) {
            'DEBUG' { 'Gray' }
            'INFO' { 'White' }
            'WARN' { 'Yellow' }
            'ERROR' { 'Red' }
        }
        Write-Host "[$Level]$functionPart $Message" -ForegroundColor $color
    }
}

function Write-FunctionEntry {
    param(
        [string]$FunctionName,
        [hashtable]$Parameters = @{}
    )
    
    $paramString = ($Parameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
    Write-Log -Level 'DEBUG' -Function $FunctionName -Message "ENTER: $paramString"
}

function Write-FunctionExit {
    param(
        [string]$FunctionName,
        [object]$Result = $null
    )
    
    $resultString = if ($Result) { "Result: $Result" } else { '' }
    Write-Log -Level 'DEBUG' -Function $FunctionName -Message "EXIT: $resultString"
}

function Test-SafePath {
    param([string]$Path)

    try {
        if (-not (Test-Path -Path $Path -PathType Leaf)) {
            return $false
        }

        if ($Path -match '\.\.|/\.\.|\\\.\.') {
            Write-Log "WARN" "Unsafe path $Path"
            return $false
        }

        return $true
    }
    catch {
        return $false
    }
}

function Split-FullName {
    param([string]$FullName)

    if ($FullName -match '^([^,]+),\s*(.+)$') {
        return @{
            FirstName = $matches[2].Trim()
            LastName  = $matches[1].Trim()
        }
    }
    elseif ($FullName -match '\s') {
        $parts = $FullName -split '\s+', 2
        return @{
            FirstName = $parts[0].Trim()
            LastName  = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' }
        }
    }
    else {
        return @{
            FirstName = $FullName.Trim()
            LastName  = ''
        }
    }
}

function Get-SearchKey {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return 'Other'
    }

    $first = $Name.Substring(0, 1).ToUpper()

    switch -Regex ($first) {
        '[A-C]' { return 'Abc' }
        '[D-F]' { return 'Def' }
        '[G-I]' { return 'Ghi' }
        '[J-L]' { return 'Jkl' }
        '[M-O]' { return 'Mno' }
        '[P-S]' { return 'Pqrs' }
        '[T-V]' { return 'Tuv' }
        '[W-Z]' { return 'Wxyz' }
        default { return 'Other' }
    }
}

function ConvertTo-NormalizedContact {
    <#
    .SYNOPSIS
        Converts a contact from any brand format to a normalized schema.
    
    .DESCRIPTION
        Creates a standard contact object with Email, FirstName, LastName, and DisplayName.
        Applies validation and sanitization during normalization.
    
    .PARAMETER Contact
        Source contact object (from CSV parsing)
    
    .PARAMETER SourceBrand
        Brand of the source format (Canon, Sharp, Xerox, Develop)
    
    .OUTPUTS
        Normalized contact hashtable: @{Email, FirstName, LastName, DisplayName}
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Contact,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
        [string]$SourceBrand
    )

    Write-FunctionEntry -FunctionName 'ConvertTo-NormalizedContact' -Parameters @{ SourceBrand = $SourceBrand }

    try {
        $config = $Script:BrandConfig[$SourceBrand]
        
        # Extract email
        $email = $Contact.($config.EmailField)
        if ([string]::IsNullOrWhiteSpace($email)) {
            Write-Log -Level 'WARN' -Function 'ConvertTo-NormalizedContact' -Message "Contact missing email address"
            return $null
        }
        
        $email = $email.Trim()
        
        # Validate email
        if (-not (Test-Email -Email $email)) {
            Write-Log -Level 'WARN' -Function 'ConvertTo-NormalizedContact' -Message "Invalid email format: $email"
            return $null
        }
        
        # Extract and parse name
        $displayName = ''
        $firstName = ''
        $lastName = ''
        
        if ($SourceBrand -eq 'Xerox' -and $config.NameFieldAlt) {
            # Xerox has separate FirstName/LastName fields
            $displayName = $Contact.($config.NameField)
            
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $firstName = $Contact.($config.NameFieldAlt[0])
                $lastName = $Contact.($config.NameFieldAlt[1])
                
                if (-not [string]::IsNullOrWhiteSpace($firstName) -or -not [string]::IsNullOrWhiteSpace($lastName)) {
                    $displayName = "$firstName $lastName".Trim()
                }
            }
            else {
                # Parse DisplayName into FirstName/LastName
                $nameParts = Split-FullName -FullName $displayName
                $firstName = $nameParts.FirstName
                $lastName = $nameParts.LastName
            }
        }
        else {
            # Other brands have single name field
            $displayName = $Contact.($config.NameField)
            
            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $displayName.Trim()
                $nameParts = Split-FullName -FullName $displayName
                $firstName = $nameParts.FirstName
                $lastName = $nameParts.LastName
            }
        }
        
        # Validate name
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            Write-Log -Level 'WARN' -Function 'ConvertTo-NormalizedContact' -Message "Contact missing name: $email"
            return $null
        }
        
        # Create normalized contact
        $normalized = @{
            Email       = $email
            FirstName   = $firstName.Trim()
            LastName    = $lastName.Trim()
            DisplayName = $displayName.Trim()
        }
        
        Write-Log -Level 'DEBUG' -Function 'ConvertTo-NormalizedContact' -Message "Normalized: $($normalized.DisplayName) <$($normalized.Email)>"
        Write-FunctionExit -FunctionName 'ConvertTo-NormalizedContact' -Result $normalized.Email
        
        return $normalized
    }
    catch {
        Write-Log -Level 'ERROR' -Function 'ConvertTo-NormalizedContact' -Message "Normalization failed" -ErrorRecord $_
        return $null
    }
}

function ConvertFrom-NormalizedContact {
    <#
    .SYNOPSIS
        Converts a normalized contact to a target brand format.
    
    .DESCRIPTION
        Maps normalized contact fields (Email, FirstName, LastName, DisplayName)
        to target brand-specific field names and values using BrandConfig OutputFields.
        Handles brand-specific transformations (e.g., name truncation for Canon,
        search keys for Develop).
    
    .PARAMETER NormalizedContact
        Normalized contact hashtable: @{Email, FirstName, LastName, DisplayName}
    
    .PARAMETER TargetBrand
        Brand of the target format (Canon, Sharp, Xerox, Develop)
    
    .OUTPUTS
        Hashtable with target brand field names and values
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$NormalizedContact,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
        [string]$TargetBrand
    )

    Write-FunctionEntry -FunctionName 'ConvertFrom-NormalizedContact' -Parameters @{ TargetBrand = $TargetBrand; Email = $NormalizedContact.Email }

    try {
        $config = $Script:BrandConfig[$TargetBrand]
        $targetContact = @{}
        
        # Map fields based on target brand
        switch ($TargetBrand) {
            'Canon' {
                $displayName = $NormalizedContact.DisplayName
                $targetContact = @{
                    cn          = $displayName
                    cnread      = $displayName
                    cnshort     = $displayName.Substring(0, [Math]::Min(13, $displayName.Length))
                    mailaddress = $NormalizedContact.Email
                }
            }
            'Sharp' {
                $targetContact = @{
                    name            = $NormalizedContact.DisplayName
                    'search-string' = $NormalizedContact.DisplayName
                    'mail-address'  = $NormalizedContact.Email
                }
            }
            'Xerox' {
                $targetContact = @{
                    DisplayName     = $NormalizedContact.DisplayName
                    FirstName       = $NormalizedContact.FirstName
                    LastName        = $NormalizedContact.LastName
                    'E-mailAddress' = $NormalizedContact.Email
                }
            }
            'Develop' {
                $targetContact = @{
                    Name        = $NormalizedContact.DisplayName
                    Furigana    = $NormalizedContact.DisplayName
                    SearchKey   = Get-SearchKey -Name $NormalizedContact.DisplayName
                    MailAddress = $NormalizedContact.Email
                }
            }
        }
        
        Write-Log -Level 'DEBUG' -Function 'ConvertFrom-NormalizedContact' -Message "Mapped to $TargetBrand format: $($NormalizedContact.DisplayName)"
        Write-FunctionExit -FunctionName 'ConvertFrom-NormalizedContact' -Result $TargetBrand
        
        return $targetContact
    }
    catch {
        Write-Log -Level 'ERROR' -Function 'ConvertFrom-NormalizedContact' -Message "Conversion to $TargetBrand failed" -ErrorRecord $_
        return $null
    }
}

#endregion

#region Parsing Functions

function Get-CsvStructure {
    <#
    .SYNOPSIS
        Extracts headers, contact data rows, and footers from CSV files.
    
    .DESCRIPTION
        Analyzes CSV files and separates:
        - Headers: All lines before the first contact (row with @)
        - Contacts: All lines containing @ symbol (email addresses)
        - Footers: All lines after the last contact
        Preserves empty lines and comments exactly as they appear.
    
    .PARAMETER FilePath
        Path to the CSV file to analyze
    
    .PARAMETER Encoding
        File encoding (UTF8, Unicode, etc.)
    
    .OUTPUTS
        Hashtable with Headers, Contacts, and Footers arrays
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$Encoding = 'UTF8'
    )

    Write-FunctionEntry -FunctionName 'Get-CsvStructure' -Parameters @{ FilePath = $FilePath; Encoding = $Encoding }

    try {
        # Read all lines preserving empty lines
        $allLines = Get-Content -Path $FilePath -Encoding $Encoding
        
        if ($null -eq $allLines -or $allLines.Count -eq 0) {
            return @{
                Headers  = @()
                Contacts = @()
                Footers  = @()
            }
        }

        # Find first and last contact row indices (rows with @ symbol)
        $firstContactIndex = -1
        $lastContactIndex = -1

        for ($i = 0; $i -lt $allLines.Count; $i++) {
            if ($allLines[$i] -match '@') {
                if ($firstContactIndex -eq -1) {
                    $firstContactIndex = $i
                }
                $lastContactIndex = $i
            }
        }

        # No contacts found - treat entire file as headers
        if ($firstContactIndex -eq -1) {
            return @{
                Headers  = $allLines
                Contacts = @()
                Footers  = @()
            }
        }

        # Extract sections
        $headers = if ($firstContactIndex -gt 0) {
            $allLines[0..($firstContactIndex - 1)]
        }
        else {
            @()
        }

        $contacts = $allLines[$firstContactIndex..$lastContactIndex]

        $footers = if ($lastContactIndex -lt ($allLines.Count - 1)) {
            $allLines[($lastContactIndex + 1)..($allLines.Count - 1)]
        }
        else {
            @()
        }

        Write-Log -Level 'INFO' -Function 'Get-CsvStructure' -Message "Extracted: $($headers.Count) headers, $($contacts.Count) contacts, $($footers.Count) footers"

        $result = @{
            Headers  = $headers
            Contacts = $contacts
            Footers  = $footers
        }
        
        Write-FunctionExit -FunctionName 'Get-CsvStructure' -Result "$($contacts.Count) contacts"
        return $result
    }
    catch {
        Write-Log -Level 'ERROR' -Function 'Get-CsvStructure' -Message "Failed for ${FilePath}" -ErrorRecord $_
        throw
    }
}

#endregion

#region Detection Functions

function Get-FileEncoding {
    param([string]$FilePath)

    try {
        $bytes = Get-Content -Path $FilePath -Encoding Byte -TotalCount 4 -ErrorAction Stop

        if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
            return 'Unicode'
        }
        elseif ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
            return 'BigEndianUnicode'
        }
        else {
            return 'UTF8'
        }
    }
    catch {
        return 'UTF8'
    }
}

function Get-SourceBrand {
    param([string]$FilePath)

    $encoding = Get-FileEncoding -FilePath $FilePath

    try {
        if ($encoding -eq 'Unicode') {
            $content = Get-Content -Path $FilePath -Encoding $encoding -TotalCount 10
            $headerLine = $content | Where-Object { $_ -match '^AbbrNo' } | Select-Object -First 1
            if ($headerLine) {
                $columns = $headerLine -split "`t"
            }
            else {
                return @{ Brand = 'Unknown'; Confidence = 0 }
            }
        }
        else {
            $allLines = Get-Content -Path $FilePath -Encoding $encoding -TotalCount 20
            $csvLines = $allLines | Where-Object { -not ($_ -match '^#') -and -not ([string]::IsNullOrWhiteSpace($_)) }

            if ($csvLines.Count -lt 1) {
                return @{ Brand = 'Unknown'; Confidence = 0 }
            }

            $headerLine = $csvLines[0]
            $columns = ($headerLine -split ',') | ForEach-Object { $_.Trim('"').Trim() }
        }

        foreach ($brand in $Script:BrandConfig.Keys) {
            $config = $Script:BrandConfig[$brand]
            $matchCount = 0

            foreach ($sigCol in $config.SignatureColumns) {
                if ($columns -contains $sigCol) {
                    $matchCount++
                }
            }

            $confidence = [math]::Round(($matchCount / $config.SignatureColumns.Count) * 100)

            if ($confidence -eq 100) {
                return @{
                    Brand      = $brand
                    Confidence = $confidence
                }
            }
        }

        return @{ Brand = 'Unknown'; Confidence = 0 }
    }
    catch {
        Write-Log "ERROR" "Detection failed for ${FilePath} $_"
        return @{ Brand = 'Unknown'; Confidence = 0 }
    }
}

#endregion

#region I/O Functions

function Read-AddressBook {
    param(
        [string]$FilePath,
        [string]$Brand
    )

    Write-FunctionEntry -FunctionName 'Read-AddressBook' -Parameters @{ FilePath = $FilePath; Brand = $Brand }

    if (-not (Test-SafePath -Path $FilePath)) {
        Write-Log -Level 'ERROR' -Function 'Read-AddressBook' -Message "Invalid file path: $FilePath"
        throw "Invalid file path"
    }

    $config = $Script:BrandConfig[$Brand]
    $contacts = @()

    try {
        # Detect encoding for the file
        $encoding = Get-FileEncoding -FilePath $FilePath
        
        # Extract CSV structure (headers, contacts, footers)
        $structure = Get-CsvStructure -FilePath $FilePath -Encoding $encoding
        
        if ($structure.Contacts.Count -eq 0) {
            Write-Log "WARN" "No contact lines found in $FilePath"
            return @()
        }

        # Parse contact lines based on brand format
        if ($Brand -eq 'Develop') {
            # Develop files need special handling - skip metadata rows and use tab delimiter
            $dataLines = $structure.Contacts
            $tempFile = [System.IO.Path]::GetTempFileName()
            $dataLines | Out-File -FilePath $tempFile -Encoding Unicode
            $data = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        elseif ($Brand -eq 'Canon') {
            # Canon: Filter out comment lines (starting with #) from contact section
            $csvLines = $structure.Contacts | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }

            if ($csvLines.Count -eq 0) {
                Write-Log "WARN" "No data lines in Canon file after filtering comments"
                return @()
            }

            $tempFile = [System.IO.Path]::GetTempFileName()
            $csvLines | Out-File -FilePath $tempFile -Encoding UTF8
            $data = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        else {
            # Sharp, Xerox: Standard CSV parsing of contact lines
            $tempFile = [System.IO.Path]::GetTempFileName()
            $structure.Contacts | Out-File -FilePath $tempFile -Encoding $encoding
            $data = Import-Csv -Path $tempFile -Encoding $encoding -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }

        # Extract email and name from each row
        foreach ($row in $data) {
            $email = $row.($config.EmailField)

            if ([string]::IsNullOrWhiteSpace($email)) {
                continue
            }

            # Handle Xerox name fields (DisplayName or FirstName+LastName)
            if ($Brand -eq 'Xerox' -and $config.NameFieldAlt) {
                $name = $row.($config.NameField)

                if ([string]::IsNullOrWhiteSpace($name)) {
                    $firstName = $row.($config.NameFieldAlt[0])
                    $lastName = $row.($config.NameFieldAlt[1])

                    if (-not [string]::IsNullOrWhiteSpace($firstName) -or -not [string]::IsNullOrWhiteSpace($lastName)) {
                        $name = "$firstName $lastName".Trim()
                    }
                }
            }
            else {
                $name = $row.($config.NameField)
            }

            # Fallback: derive name from email if missing
            if ([string]::IsNullOrWhiteSpace($name)) {
                $name = ($email -split '@')[0]
            }

            $contacts += [PSCustomObject]@{
                Name  = $name.Trim()
                Email = $email.Trim()
            }
        }

        Write-Log -Level 'INFO' -Function 'Read-AddressBook' -Message "Read $($contacts.Count) contacts from $FilePath"
        Write-FunctionExit -FunctionName 'Read-AddressBook' -Result "$($contacts.Count) contacts"
        return $contacts
    }
    catch {
        Write-Log -Level 'ERROR' -Function 'Read-AddressBook' -Message "Read failed for ${FilePath}" -ErrorRecord $_
        throw
    }
}

function Write-AddressBook {
    <#
    .SYNOPSIS
        Writes normalized contacts to target brand CSV format.
    
    .DESCRIPTION
        Accepts normalized contact objects and converts them to target brand format.
        Uses ConvertFrom-NormalizedContact for field mapping and fills all required
        brand-specific fields with defaults. Preserves encoding and delimiters.
    
    .PARAMETER NormalizedContacts
        Array of normalized contact hashtables: @{Email, FirstName, LastName, DisplayName}
    
    .PARAMETER OutputPath
        Path where the output CSV will be written
    
    .PARAMETER TargetBrand
        Brand format for output (Canon, Sharp, Xerox, Develop)
    
    .OUTPUTS
        None - writes file to OutputPath
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$NormalizedContacts,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
        [string]$TargetBrand
    )

    Write-FunctionEntry -FunctionName 'Write-AddressBook' -Parameters @{ OutputPath = $OutputPath; TargetBrand = $TargetBrand; ContactCount = $NormalizedContacts.Count }

    try {
        $output = @()
        $id = 1

        # Convert each normalized contact to target brand format
        foreach ($normalized in $NormalizedContacts) {
            # Map to target brand fields
            $targetFields = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TargetBrand $TargetBrand
            
            if ($null -eq $targetFields) {
                Write-Log -Level 'WARN' -Function 'Write-AddressBook' -Message "Skipping contact - conversion failed: $($normalized.DisplayName)"
                continue
            }

            # Build complete brand-specific object with all required fields
            switch ($TargetBrand) {
                'Canon' {
                    $output += [PSCustomObject]@{
                        objectclass       = 'email'
                        cn                = $targetFields.cn
                        cnread            = $targetFields.cnread
                        cnshort           = $targetFields.cnshort
                        subdbid           = 11
                        mailaddress       = $targetFields.mailaddress
                        dialdata          = ''
                        uri               = ''
                        url               = ''
                        path              = ''
                        protocol          = 'smtp'
                        username          = ''
                        pwd               = ''
                        member            = ''
                        indxid            = $id
                        enablepartial     = 'off'
                        sub               = ''
                        faxprotocol       = ''
                        ecm               = ''
                        txstartspeed      = ''
                        commode           = ''
                        lineselect        = ''
                        uricommode        = ''
                        uriflag           = ''
                        pwdinputflag      = ''
                        ifaxmode          = ''
                        transsvcstr1      = ''
                        transsvcstr2      = ''
                        ifaxdirectmode    = ''
                        documenttype      = ''
                        bwpapersize       = ''
                        bwcompressiontype = ''
                        bwpixeltype       = ''
                        bwbitsperpixel    = ''
                        bwresolution      = ''
                        clpapersize       = ''
                        clcompressiontype = ''
                        clpixeltype       = ''
                        clbitsperpixel    = ''
                        clresolution      = ''
                        accesscode        = ''
                        uuid              = ''
                        cnreadlang        = 'en'
                        enablesfp         = ''
                        memberobjectuuid  = ''
                        loginusername     = ''
                        logindomainname   = ''
                        usergroupname     = ''
                        personalid        = ''
                        folderidflag      = ''
                    }
                }
                'Sharp' {
                    $output += [PSCustomObject]@{
                        address                            = 'data'
                        'search-id'                        = $id
                        name                               = $targetFields.name
                        'search-string'                    = $targetFields.'search-string'
                        'category-id'                      = 1
                        'frequently-used'                  = 'FALSE'
                        'mail-address'                     = $targetFields.'mail-address'
                        'fax-number'                       = ''
                        'ifax-address'                     = ''
                        'ftp-host'                         = ''
                        'ftp-directory'                    = ''
                        'ftp-username'                     = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'ftp-username/@encodingMethod'     = 'encrypted2'
                        'ftp-password'                     = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'ftp-password/@encodingMethod'     = 'encrypted2'
                        'smb-directory'                    = ''
                        'smb-username'                     = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'smb-username/@encodingMethod'     = 'encrypted2'
                        'smb-password'                     = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'smb-password/@encodingMethod'     = 'encrypted2'
                        'desktop-host'                     = ''
                        'desktop-port'                     = ''
                        'desktop-directory'                = ''
                        'desktop-username'                 = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'desktop-username/@encodingMethod' = 'encrypted2'
                        'desktop-password'                 = '+xS4FiNvCE4i8EqfPNhjWg=='
                        'desktop-password/@encodingMethod' = 'encrypted2'
                    }
                }
                'Xerox' {
                    $output += [PSCustomObject]@{
                        XrxAddressBookId       = $id
                        DisplayName            = $targetFields.DisplayName
                        FirstName              = $targetFields.FirstName
                        LastName               = $targetFields.LastName
                        Company                = ''
                        XrxAllFavoritesOrder   = ''
                        MemberOf               = '""""""'
                        IsDL                   = 0
                        XrxApplicableWorkflows = ''
                        FaxNumber              = ''
                        XrxIsFaxFavorite       = 0
                        'E-mailAddress'        = $targetFields.'E-mailAddress'
                        XrxIsEmailFavorite     = 0
                        InternetFaxAddress     = ''
                        ScanNickName           = ''
                        XrxIsScanFavorite      = 0
                        ScanTransferProtocol   = 4
                        ScanServerAddress      = '(null)'
                        ScanServerPort         = 0
                        ScanDocumentPath       = ''
                        ScanLoginName          = ''
                        ScanLoginPassword      = ''
                        ScanSMBShare           = ''
                        ScanNDSTree            = ''
                        ScanNDSContext         = ''
                        ScanNDSVolume          = ''
                    }
                }
                'Develop' {
                    $output += [PSCustomObject]@{
                        AbbrNo              = $id
                        Name                = $targetFields.Name
                        Pinyin              = 'No'
                        Furigana            = $targetFields.Furigana
                        SearchKey           = $targetFields.SearchKey
                        WellUse             = 'Yes'
                        SendMode            = 'Email'
                        IconID              = ''
                        UseReferLicence     = 'Level'
                        ReferGroupNo        = 0
                        ReferPossibleLevel  = 0
                        MailAddress         = $targetFields.MailAddress
                        FTPServerAddress    = ''
                        FTPServerFolder     = ''
                        FTPLoginAnonymous   = ''
                        FTPLoginUser        = ''
                        FTPLoginPassword    = ''
                        FTPPassiveSend      = ''
                        FTPProxy            = ''
                        FTPPortNo           = ''
                        SMBAddress          = ''
                        SMBFolder           = ''
                        SMBLoginUser        = ''
                        SMBLoginPassword    = ''
                        WebDAVServerAddress = ''
                        WebDAVCollection    = ''
                        WebDAVLoginUser     = ''
                        WebDAVLoginPassword = ''
                        WebDAVSSL           = ''
                        WebDAVProxy         = ''
                        WebDAVPortNo        = ''
                        BoxID               = ''
                        Model               = ''
                        FaxPhoneNo          = ''
                        FaxCapability       = ''
                        FaxV34Off           = ''
                        FaxECMOff           = ''
                        FaxOversea          = ''
                        FaxLine             = ''
                        CheckDest           = ''
                        Host                = ''
                        PortNo              = ''
                        IfaxResolution      = ''
                        IfaxSize            = ''
                        IfaxCompression     = ''
                    }
                }
            }
            $id++
        }

        # Write output with brand-specific encoding and format
        if ($TargetBrand -eq 'Develop') {
            $timestamp = Get-Date -Format 'yyyy.M.d HH:mm:ss'
            $header1 = "@Ver406`t22C-6e`tIntegrate`tUTF-16LE`t${timestamp}`tabbr`t000ac209552c58a5b6222cf539ab712a255b`t"
            $header2 = "#Abbreviate`t2000`t"

            $tempFile = [System.IO.Path]::GetTempFileName()
            $output | Export-Csv -Path $tempFile -Delimiter "`t" -NoTypeInformation -Encoding UTF8

            $csvContent = Get-Content -Path $tempFile -Encoding UTF8
            @($header1, $header2) + $csvContent | Out-File -FilePath $OutputPath -Encoding Unicode

            Remove-Item -Path $tempFile -Force
        }
        elseif ($TargetBrand -eq 'Canon') {
            # Canon header block
            $header = @"
`# Canon AddressBook CSV version: 0x0003
# CharSet: UTF-8
# dn: fixed
# DB Version: 0x010b

"@

            # Canon CSV column header - UNQUOTED with trailing comma
            $canonColumns = "objectclass,cn,cnread,cnshort,subdbid,mailaddress,dialdata,uri,url,path,protocol,username,pwd,member,indxid,enablepartial,sub,faxprotocol,ecm,txstartspeed,commode,lineselect,uricommode,uriflag,pwdinputflag,ifaxmode,transsvcstr1,transsvcstr2,ifaxdirectmode,documenttype,bwpapersize,bwcompressiontype,bwpixeltype,bwbitsperpixel,bwresolution,clpapersize,clcompressiontype,clpixeltype,clbitsperpixel,clresolution,accesscode,uuid,cnreadlang,enablesfp,memberobjectuuid,loginusername,logindomainname,usergroupname,personalid,folderidflag,"

            # Write header block + CSV column header
            [System.IO.File]::WriteAllText($OutputPath, $header + $canonColumns + "`r`n", [System.Text.Encoding]::UTF8)

            # Write data rows
            foreach ($obj in $output) {
                $row = @(
                    $obj.objectclass,
                    "`"$($obj.cn)`"",
                    "`"$($obj.cnread)`"",
                    "`"$($obj.cnshort)`"",
                    $obj.subdbid,
                    "`"$($obj.mailaddress)`"",
                    $obj.dialdata, $obj.uri, $obj.url, $obj.path, $obj.protocol,
                    $obj.username, $obj.pwd, $obj.member, $obj.indxid, $obj.enablepartial,
                    $obj.sub, $obj.faxprotocol, $obj.ecm, $obj.txstartspeed, $obj.commode,
                    $obj.lineselect, $obj.uricommode, $obj.uriflag, $obj.pwdinputflag,
                    $obj.ifaxmode, $obj.transsvcstr1, $obj.transsvcstr2, $obj.ifaxdirectmode,
                    $obj.documenttype, $obj.bwpapersize, $obj.bwcompressiontype, $obj.bwpixeltype,
                    $obj.bwbitsperpixel, $obj.bwresolution, $obj.clpapersize, $obj.clcompressiontype,
                    $obj.clpixeltype, $obj.clbitsperpixel, $obj.clresolution, $obj.accesscode,
                    $obj.uuid, $obj.cnreadlang, $obj.enablesfp, $obj.memberobjectuuid,
                    $obj.loginusername, $obj.logindomainname, $obj.usergroupname, $obj.personalid,
                    $obj.folderidflag
                )

                $line = $row -join ','
                [System.IO.File]::AppendAllText($OutputPath, $line + "`r`n", [System.Text.Encoding]::UTF8)
            }
        }
        else {
            # Sharp and Xerox use standard CSV export
            $output | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        }

        Write-Log -Level 'INFO' -Function 'Write-AddressBook' -Message "Wrote $($output.Count) contacts to $OutputPath ($TargetBrand)"
        Write-FunctionExit -FunctionName 'Write-AddressBook' -Result "$($output.Count) contacts written"
    }
    catch {
        Write-Log -Level 'ERROR' -Function 'Write-AddressBook' -Message "Write failed to ${OutputPath}" -ErrorRecord $_
        throw
    }
}

#endregion

#region Validation Functions

function Test-Email {
    param([string]$Email)

    if ([string]::IsNullOrWhiteSpace($Email)) {
        return $false
    }

    $pattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return $Email -match $pattern
}

function Remove-Duplicates {
    param([array]$Contacts)

    $seen = @{}
    $unique = @()
    $duplicates = 0

    foreach ($contact in $Contacts) {
        $emailLower = $contact.Email.ToLower()

        if ($seen.ContainsKey($emailLower)) {
            $duplicates++
        }
        else {
            $seen[$emailLower] = $true
            $unique += $contact
        }
    }

    if ($duplicates -gt 0) {
        $Script:Stats.Duplicates += $duplicates
        Write-Log "INFO" "Removed $duplicates duplicates"
    }

    return $unique
}

function Validate-Contacts {
    param([array]$Contacts)

    $valid = @()

    foreach ($contact in $Contacts) {
        if (Test-Email -Email $contact.Email) {
            $valid += $contact
            $Script:Stats.Converted++
        }
        else {
            $Script:Stats.Skipped++
            $Script:Stats.InvalidEmails += $contact.Email
            Write-Log "WARN" "Skipped invalid $($contact.Email)"
        }
    }

    return $valid
}

function Validate-OutputFile {
    <#
    .SYNOPSIS
        Validates output CSV file structure and content.
    
    .DESCRIPTION
        Performs quality checks on output CSV:
        - File exists and is readable
        - Has correct structure for target brand
        - Email/name fields are populated in correct columns
        - No corrupted lines or column misalignment
        - Encoding matches target brand requirements
    
    .PARAMETER FilePath
        Path to output CSV file to validate
    
    .PARAMETER TargetBrand
        Brand of the target format (Canon, Sharp, Xerox, Develop)
    
    .PARAMETER ExpectedContactCount
        Expected number of contacts in output (optional)
    
    .OUTPUTS
        Hashtable with validation results: @{IsValid, Errors[], Warnings[], ContactCount}
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
        [string]$TargetBrand,
        
        [Parameter(Mandatory = $false)]
        [int]$ExpectedContactCount = -1
    )

    Write-FunctionEntry -FunctionName 'Validate-OutputFile' -Parameters @{ FilePath = $FilePath; TargetBrand = $TargetBrand }

    $result = @{
        IsValid      = $true
        Errors       = @()
        Warnings     = @()
        ContactCount = 0
    }

    try {
        # Check file exists
        if (-not (Test-Path $FilePath)) {
            $result.Errors += "File does not exist: $FilePath"
            $result.IsValid = $false
            Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "File not found: $FilePath"
            return $result
        }

        # Check file is not empty
        $fileInfo = Get-Item $FilePath
        if ($fileInfo.Length -eq 0) {
            $result.Errors += "File is empty: $FilePath"
            $result.IsValid = $false
            Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "File is empty: $FilePath"
            return $result
        }

        # Get expected encoding for brand
        $expectedEncoding = switch ($TargetBrand) {
            'Develop' { 'Unicode' }
            default { 'UTF8' }
        }

        # Read file and check structure
        try {
            $structure = Get-CsvStructure -FilePath $FilePath -Encoding $expectedEncoding
            $result.ContactCount = $structure.Contacts.Count
            
            Write-Log -Level 'INFO' -Function 'Validate-OutputFile' -Message "Found $($structure.Contacts.Count) contact rows"
            
            if ($structure.Contacts.Count -eq 0) {
                $result.Warnings += "No contact rows found in output file"
                Write-Log -Level 'WARN' -Function 'Validate-OutputFile' -Message "No contacts in output"
            }
        }
        catch {
            $result.Errors += "Failed to parse CSV structure: $_"
            $result.IsValid = $false
            Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "CSV parsing failed" -ErrorRecord $_
            return $result
        }

        # Check expected contact count if provided
        if ($ExpectedContactCount -gt 0 -and $result.ContactCount -ne $ExpectedContactCount) {
            $result.Warnings += "Contact count mismatch: Expected $ExpectedContactCount, found $($result.ContactCount)"
            Write-Log -Level 'WARN' -Function 'Validate-OutputFile' -Message "Count mismatch: expected $ExpectedContactCount, got $($result.ContactCount)"
        }

        # Brand-specific validation
        $config = $Script:BrandConfig[$TargetBrand]
        
        # Parse contacts to check field structure
        try {
            if ($TargetBrand -eq 'Develop') {
                $tempFile = [System.IO.Path]::GetTempFileName()
                $structure.Contacts | Out-File -FilePath $tempFile -Encoding Unicode
                $data = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
                Remove-Item -Path $tempFile -Force
            }
            elseif ($TargetBrand -eq 'Canon') {
                # Skip comment lines for Canon
                $csvLines = $structure.Contacts | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }
                if ($csvLines.Count -gt 0) {
                    $tempFile = [System.IO.Path]::GetTempFileName()
                    $csvLines | Out-File -FilePath $tempFile -Encoding UTF8
                    $data = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
                    Remove-Item -Path $tempFile -Force
                }
            }
            else {
                $tempFile = [System.IO.Path]::GetTempFileName()
                $structure.Contacts | Out-File -FilePath $tempFile -Encoding $expectedEncoding
                $data = Import-Csv -Path $tempFile -Encoding $expectedEncoding -Delimiter $config.Delimiter
                Remove-Item -Path $tempFile -Force
            }

            # Check for required fields
            if ($data.Count -gt 0) {
                $firstRow = $data[0]
                $emailField = $config.OutputFields.Email
                
                if (-not $firstRow.PSObject.Properties[$emailField]) {
                    $result.Errors += "Missing required email field: $emailField"
                    $result.IsValid = $false
                    Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "Missing email field: $emailField"
                }

                # Check if email fields have values
                $emptyEmails = 0
                foreach ($row in $data) {
                    if ([string]::IsNullOrWhiteSpace($row.($emailField))) {
                        $emptyEmails++
                    }
                }

                if ($emptyEmails -gt 0) {
                    $result.Warnings += "$emptyEmails contacts have empty email addresses"
                    Write-Log -Level 'WARN' -Function 'Validate-OutputFile' -Message "$emptyEmails empty emails"
                }
            }
        }
        catch {
            $result.Errors += "Failed to validate contact data: $_"
            $result.IsValid = $false
            Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "Contact validation failed" -ErrorRecord $_
            return $result
        }

        Write-Log -Level 'INFO' -Function 'Validate-OutputFile' -Message "Validation complete: IsValid=$($result.IsValid), Contacts=$($result.ContactCount)"
        Write-FunctionExit -FunctionName 'Validate-OutputFile' -Result "IsValid=$($result.IsValid)"
        
        return $result
    }
    catch {
        $result.Errors += "Validation failed: $_"
        $result.IsValid = $false
        Write-Log -Level 'ERROR' -Function 'Validate-OutputFile' -Message "Validation error" -ErrorRecord $_
        return $result
    }
}

#endregion

#region UI Functions

function Show-Menu {
    param(
        [string[]]$Options,
        [string]$Title,
        [string]$Description = ""
    )

    $selected = 0

    while ($true) {
        Clear-Host

        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host "  PRINTER ADDRESS BOOK CONVERTER v1.5" -ForegroundColor Cyan
        Write-Host "  Canon | Sharp | Xerox | Develop/Konica" -ForegroundColor Cyan
        Write-Host "  by Faris Khasawneh - January 2026" -ForegroundColor Gray
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""

        if ($Description) {
            Write-Host "$Description" -ForegroundColor Gray
            Write-Host ""
        }

        Write-Host "$Title" -ForegroundColor White
        Write-Host ""

        for ($i = 0; $i -lt $Options.Length; $i++) {
            if ($i -eq $selected) {
                Write-Host "  > " -NoNewline -ForegroundColor Green
                Write-Host $Options[$i] -ForegroundColor Green
            }
            else {
                Write-Host "    $($Options[$i])"
            }
        }

        Write-Host ""
        Write-Host "Use up/down arrows, Enter to select, ESC to cancel" -ForegroundColor Gray

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        switch ($key.VirtualKeyCode) {
            38 { $selected = [Math]::Max(0, $selected - 1) }
            40 { $selected = [Math]::Min($Options.Length - 1, $selected + 1) }
            13 { return $selected }
            27 { return -1 }
        }
    }
}

function Select-Files {
    param([bool]$MultiSelect = $false)

    try {
        Add-Type -AssemblyName System.Windows.Forms

        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.Filter = "CSV/TSV files (*.csv;*.tsv)|*.csv;*.tsv|All files (*.*)|*.*"
        $FileBrowser.Title = "Select Address Book File(s)"
        $FileBrowser.MultiSelect = $MultiSelect

        if ($FileBrowser.ShowDialog() -eq 'OK') {
            if ($MultiSelect) {
                return $FileBrowser.FileNames
            }
            else {
                return $FileBrowser.FileName
            }
        }

        return $null
    }
    catch {
        Write-Host "Error: File picker failed" -ForegroundColor Red
        return $null
    }
}

function Show-ValidationReport {
    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "  CONVERSION SUMMARY" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan

    Write-Host "  Converted: " -NoNewline -ForegroundColor Green
    Write-Host $Script:Stats.Converted

    if ($Script:Stats.Skipped -gt 0) {
        Write-Host "  Skipped: " -NoNewline -ForegroundColor Yellow
        Write-Host $Script:Stats.Skipped

        if ($Script:Stats.InvalidEmails.Count -gt 0) {
            $display = $Script:Stats.InvalidEmails | Select-Object -First 3
            Write-Host "    Invalid: " -NoNewline -ForegroundColor Gray
            Write-Host ($display -join ', ') -ForegroundColor Gray
            if ($Script:Stats.InvalidEmails.Count -gt 3) {
                Write-Host "    ... and $($Script:Stats.InvalidEmails.Count - 3) more" -ForegroundColor Gray
            }
        }
    }

    if ($Script:Stats.Duplicates -gt 0) {
        Write-Host "  Duplicates removed: " -NoNewline -ForegroundColor Cyan
        Write-Host $Script:Stats.Duplicates
    }

    Write-Host "  Log: " -NoNewline -ForegroundColor White
    Write-Host $Script:LogFile -ForegroundColor Gray

    Write-Host "===============================================================" -ForegroundColor Cyan
}

function Reset-Stats {
    $Script:Stats = @{
        Converted     = 0
        Skipped       = 0
        Duplicates    = 0
        InvalidEmails = @()
    }
}

function Backup-SourceFile {
    param([string]$SourcePath)

    try {
        $sourceDir = Split-Path -Parent $SourcePath
        $backupDir = Join-Path -Path $sourceDir -ChildPath 'backup'
        if (-not (Test-Path $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        }

        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
        $extension = [System.IO.Path]::GetExtension($SourcePath)
        $timestamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
        $backupPath = Join-Path -Path $backupDir -ChildPath "${fileName}_backup_${timestamp}${extension}"

        Copy-Item -Path $SourcePath -Destination $backupPath -Force
        Write-Log "INFO" "Backup $backupPath"

        Write-Host "  Backup: " -NoNewline
        Write-Host "$backupPath" -ForegroundColor Cyan
    }
    catch {
        Write-Host "  Warning: Backup failed" -ForegroundColor Yellow
    }
}

#endregion

#region Conversion Functions

function Get-SafeOutputPath {
    param(
        [string]$SourcePath,
        [string]$TargetBrand,
        [bool]$IsMerge = $false
    )

    try {
        $sourceDir = Split-Path -Parent $SourcePath
        $convertedDir = Join-Path -Path $sourceDir -ChildPath 'converted'

        if (-not (Test-Path $convertedDir)) {
            New-Item -Path $convertedDir -ItemType Directory -Force | Out-Null
        }

        if ($IsMerge) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd'
            $fileName = "Merged_to_${TargetBrand}_${timestamp}.csv"
        }
        else {
            $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
            $sourceFileName = $sourceFileName -replace '[\\/:*?"<>|]', '_'
            $fileName = "${sourceFileName}_to_${TargetBrand}.csv"
        }

        return Join-Path -Path $convertedDir -ChildPath $fileName
    }
    catch {
        throw "Failed to generate output path $_"
    }
}

function Convert-AddressBook {
    <#
    .SYNOPSIS
        Converts an address book from source brand to target brand.
    
    .DESCRIPTION
        Full conversion pipeline:
        1. Parse source CSV and extract contact rows
        2. ConvertTo-NormalizedContact: Map each row to normalized schema
        3. Validate-Contacts: Filter invalid emails
        4. Remove-Duplicates: Deduplicate by email
        5. Write-AddressBook: Write target CSV via ConvertFrom-NormalizedContact
    
    .PARAMETER SourcePath
        Path to source CSV file
    
    .PARAMETER SourceBrand
        Brand of source format
    
    .PARAMETER TargetBrand
        Brand of target format
    
    .OUTPUTS
        Array of processed normalized contacts
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceBrand,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetBrand
    )

    Write-FunctionEntry -FunctionName 'Convert-AddressBook' -Parameters @{ SourcePath = $SourcePath; SourceBrand = $SourceBrand; TargetBrand = $TargetBrand }

    Write-Host ""
    Write-Host "Processing: " -NoNewline
    Write-Host (Split-Path -Leaf $SourcePath) -ForegroundColor Cyan
    Write-Host "  Source: " -NoNewline
    Write-Host $SourceBrand -ForegroundColor Green
    Write-Host "  Target: " -NoNewline
    Write-Host $TargetBrand -ForegroundColor Green

    Backup-SourceFile -SourcePath $SourcePath

    # Step 1: Parse source CSV and get raw contact rows
    Write-Host "  Reading..." -NoNewline
    try {
        $config = $Script:BrandConfig[$SourceBrand]
        $encoding = Get-FileEncoding -FilePath $SourcePath
        $structure = Get-CsvStructure -FilePath $SourcePath -Encoding $encoding
        
        if ($structure.Contacts.Count -eq 0) {
            Write-Host " No contacts found" -ForegroundColor Yellow
            Write-Log -Level 'WARN' -Function 'Convert-AddressBook' -Message "No contact rows in source file"
            return $null
        }
        
        # Parse contact rows based on brand format
        if ($SourceBrand -eq 'Develop') {
            $tempFile = [System.IO.Path]::GetTempFileName()
            $structure.Contacts | Out-File -FilePath $tempFile -Encoding Unicode
            $sourceRows = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        elseif ($SourceBrand -eq 'Canon') {
            # Filter out comment lines
            $csvLines = $structure.Contacts | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }
            if ($csvLines.Count -eq 0) {
                Write-Host " No data rows" -ForegroundColor Yellow
                return $null
            }
            $tempFile = [System.IO.Path]::GetTempFileName()
            $csvLines | Out-File -FilePath $tempFile -Encoding UTF8
            $sourceRows = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        else {
            # Sharp, Xerox
            $tempFile = [System.IO.Path]::GetTempFileName()
            $structure.Contacts | Out-File -FilePath $tempFile -Encoding $encoding
            $sourceRows = Import-Csv -Path $tempFile -Encoding $encoding -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        
        Write-Host " $($sourceRows.Count) contacts" -ForegroundColor Green
        Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Read $($sourceRows.Count) contact rows from $SourceBrand"
    }
    catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        Write-Log -Level 'ERROR' -Function 'Convert-AddressBook' -Message "Read failed" -ErrorRecord $_
        return $null
    }

    # Step 2: Normalize contacts (convert raw CSV rows to brand-agnostic format)
    Write-Host "  Normalizing..." -NoNewline
    $normalizedContacts = @()
    foreach ($row in $sourceRows) {
        $normalized = ConvertTo-NormalizedContact -Contact $row -SourceBrand $SourceBrand
        if ($normalized) {
            $normalizedContacts += $normalized
        }
    }
    Write-Host " $($normalizedContacts.Count) normalized" -ForegroundColor Green
    Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Normalized $($normalizedContacts.Count) contacts"

    if ($normalizedContacts.Count -eq 0) {
        Write-Host "  No valid contacts to process" -ForegroundColor Yellow
        Write-Log -Level 'WARN' -Function 'Convert-AddressBook' -Message "No contacts after normalization"
        return $null
    }

    # Step 3: Validate (checks email format - already done in normalization, but track stats)
    Write-Host "  Validating..." -NoNewline
    $validCount = 0
    $validContacts = @()
    foreach ($normalized in $normalizedContacts) {
        if (Test-Email -Email $normalized.Email) {
            $validContacts += $normalized
            $validCount++
            $Script:Stats.Converted++
        }
        else {
            $Script:Stats.Skipped++
            $Script:Stats.InvalidEmails += $normalized.Email
            Write-Log -Level 'WARN' -Function 'Convert-AddressBook' -Message "Skipped invalid email: $($normalized.Email)"
        }
    }
    Write-Host " $validCount valid" -ForegroundColor Green
    Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Validated $validCount contacts"

    # Step 4: Deduplicate
    Write-Host "  Deduplicating..." -NoNewline
    $seen = @{}
    $uniqueContacts = @()
    $duplicates = 0
    foreach ($contact in $validContacts) {
        $emailLower = $contact.Email.ToLower()
        if ($seen.ContainsKey($emailLower)) {
            $duplicates++
            $Script:Stats.Duplicates++
        }
        else {
            $seen[$emailLower] = $true
            $uniqueContacts += $contact
        }
    }
    Write-Host " $($uniqueContacts.Count) unique" -ForegroundColor Green
    Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Removed $duplicates duplicates"

    # Step 5: Get output path
    try {
        $outputPath = Get-SafeOutputPath -SourcePath $SourcePath -TargetBrand $TargetBrand
        Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Output path: $outputPath"
    }
    catch {
        Write-Host "  Output path error: $_" -ForegroundColor Red
        Write-Log -Level 'ERROR' -Function 'Convert-AddressBook' -Message "Output path error" -ErrorRecord $_
        return $null
    }

    # Step 6: Write to target format
    Write-Host "  Writing..." -NoNewline
    try {
        Write-AddressBook -NormalizedContacts $uniqueContacts -OutputPath $outputPath -TargetBrand $TargetBrand
        Write-Host " Done" -ForegroundColor Green

        Write-Host "  Output: " -NoNewline -ForegroundColor White
        Write-Host $outputPath -ForegroundColor Cyan
        
        Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Conversion complete: $outputPath"
        Write-FunctionExit -FunctionName 'Convert-AddressBook' -Result "$($uniqueContacts.Count) contacts written"
    }
    catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        Write-Log -Level 'ERROR' -Function 'Convert-AddressBook' -Message "Write failed" -ErrorRecord $_
        return $null
    }

    return $uniqueContacts
}

#endregion

#region Main Execution

function Main {
    Clear-Host

    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "  Printer Address Book Converter v1.5" -ForegroundColor Cyan
    Write-Host "  Log file: $Script:LogFile" -ForegroundColor Gray
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-FunctionEntry -FunctionName 'Main' -Parameters @{ SourcePath = $SourcePath; TargetBrand = $TargetBrand; Mode = $Mode; NoInteractive = $NoInteractive }
    Write-Log -Level 'INFO' -Function 'Main' -Message "Session started"

    # Non-interactive mode
    if ($NoInteractive -and $SourcePath -and $TargetBrand) {
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host "  NON-INTERACTIVE MODE" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""

        if (-not (Test-Path $SourcePath)) {
            Write-Host "Error: Source file not found: $SourcePath" -ForegroundColor Red
            Write-Host ""
            return
        }

        if (-not (Test-SafePath -Path $SourcePath)) {
            Write-Host "Error: Invalid file path: $SourcePath" -ForegroundColor Red
            Write-Host ""
            return
        }

        # Detect source brand
        $detection = Get-SourceBrand -FilePath $SourcePath
        if ($detection.Confidence -ne 100) {
            Write-Host "Error: Could not auto-detect source brand for: $SourcePath" -ForegroundColor Red
            Write-Host ""
            return
        }

        $sourceBrand = $detection.Brand
        Write-Host "Detected source: $sourceBrand" -ForegroundColor Green
        Write-Host "Target brand: $TargetBrand" -ForegroundColor Green
        Write-Host ""

        Reset-Stats

        $converted = Convert-AddressBook -SourcePath $SourcePath -SourceBrand $sourceBrand -TargetBrand $TargetBrand

        if ($converted) {
            Show-ValidationReport
        }
        else {
            Write-Host "Conversion failed" -ForegroundColor Red
        }

        Write-Log "INFO" "Session completed (non-interactive)"
        Write-Host ""
        return
    }

    # Interactive mode (original behavior)
    $modeOptions = @(
        'Single File Conversion',
        'Batch Convert Multiple Files',
        'Merge Multiple Files into One',
        'Exit'
    )
    $desc = "Auto-detects source format and converts to target brand.`nOutput saved to 'converted' folder with automatic backups."
    $modeIndex = Show-Menu -Options $modeOptions -Title "Select conversion mode:" -Description $desc

    if ($modeIndex -eq -1 -or $modeIndex -eq 3) {
        Write-Host ""
        Write-Host "Cancelled" -ForegroundColor Yellow
        Write-Host ""
        return
    }

    $mode = switch ($modeIndex) {
        0 { 'Single' }
        1 { 'Batch' }
        2 { 'Merge' }
    }

    $sourceFiles = @()

    if ($mode -eq 'Single') {
        $sourceFile = Select-Files -MultiSelect $false
        if (-not $sourceFile) {
            Write-Host ""
            Write-Host "No file selected" -ForegroundColor Yellow
            Write-Host ""
            return
        }
        $sourceFiles += $sourceFile
    }
    else {
        $selectedFiles = Select-Files -MultiSelect $true
        if (-not $selectedFiles -or $selectedFiles.Count -eq 0) {
            Write-Host ""
            Write-Host "No files selected" -ForegroundColor Yellow
            Write-Host ""
            return
        }
        $sourceFiles = $selectedFiles
    }

    foreach ($file in $sourceFiles) {
        if (-not (Test-SafePath -Path $file)) {
            Write-Host ""
            Write-Host "Error: Invalid file path $file" -ForegroundColor Red
            Write-Host ""
            return
        }
    }

    Write-Host ""
    Write-Host "Detecting formats..." -ForegroundColor Cyan
    Write-Host ""
    $fileInfo = @()
    $failedFiles = @()

    foreach ($file in $sourceFiles) {
        $detection = Get-SourceBrand -FilePath $file
        $fileName = Split-Path -Leaf $file

        if ($detection.Confidence -eq 100) {
            Write-Host "  $fileName" -NoNewline
            Write-Host " -> " -NoNewline -ForegroundColor Gray
            Write-Host $detection.Brand -ForegroundColor Green

            $fileInfo += @{
                Path     = $file
                Brand    = $detection.Brand
                FileName = $fileName
            }
        }
        else {
            Write-Host "  $fileName" -NoNewline
            Write-Host " -> Detection failed" -ForegroundColor Yellow

            $failedFiles += @{
                Path     = $file
                FileName = $fileName
            }
        }

        Write-Log "INFO" "Detected $($detection.Brand) for $fileName $($detection.Confidence)"
    }

    if ($failedFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "Manual brand selection required for $($failedFiles.Count) file(s)" -ForegroundColor Yellow
        Write-Host ""

        foreach ($failedFile in $failedFiles) {
            $brandOptions = @('Canon', 'Sharp', 'Xerox', 'Develop', 'Skip this file')
            $brandIndex = Show-Menu -Options $brandOptions -Title "Select brand for: $($failedFile.FileName)" -Description "Auto-detection failed. Please select the source brand manually."

            if ($brandIndex -eq -1 -or $brandIndex -eq 4) {
                Write-Host "  Skipping $($failedFile.FileName)" -ForegroundColor Yellow
                Write-Log "INFO" "Skipped $($failedFile.FileName) (user cancelled)"
                continue
            }

            $selectedBrand = $brandOptions[$brandIndex]

            $fileInfo += @{
                Path     = $failedFile.Path
                Brand    = $selectedBrand
                FileName = $failedFile.FileName
            }

            Write-Host "  $($failedFile.FileName)" -NoNewline
            Write-Host " -> " -NoNewline -ForegroundColor Gray
            Write-Host "$selectedBrand (manual)" -ForegroundColor Cyan
            Write-Log "INFO" "Manual selection $selectedBrand for $($failedFile.FileName)"
        }
    }

    if ($fileInfo.Count -eq 0) {
        Write-Host ""
        Write-Host "No files to convert" -ForegroundColor Red
        Write-Host ""
        return
    }

    Write-Host ""
    $targetBrandOptions = @('Canon', 'Sharp', 'Xerox', 'Develop')
    $targetBrandIndex = Show-Menu -Options $targetBrandOptions -Title "Select target brand:" -Description "All files will be converted to this format."

    if ($targetBrandIndex -eq -1) {
        Write-Host ""
        Write-Host "Cancelled" -ForegroundColor Yellow
        Write-Host ""
        return
    }

    $targetBrand = $targetBrandOptions[$targetBrandIndex]

    Write-Log "INFO" "Target $targetBrand"

    Reset-Stats

    $allConvertedContacts = @()

    if ($mode -eq 'Merge') {
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host "  MERGE MODE" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan

        foreach ($info in $fileInfo) {
            try {
                $contacts = Read-AddressBook -FilePath $info.Path -Brand $info.Brand
                $validContacts = Validate-Contacts -Contacts $contacts
                $allConvertedContacts += $validContacts

                Write-Host "  Loaded $($validContacts.Count) from " -NoNewline
                Write-Host $info.FileName -ForegroundColor Cyan
            }
            catch {
                Write-Host "  Failed: " -NoNewline -ForegroundColor Red
                Write-Host $info.FileName -ForegroundColor Red
            }
        }

        Write-Host ""
        Write-Host "  Total: " -NoNewline
        Write-Host $allConvertedContacts.Count -ForegroundColor Green

        $uniqueContacts = Remove-Duplicates -Contacts $allConvertedContacts
        Write-Host "  Unique: " -NoNewline
        Write-Host $uniqueContacts.Count -ForegroundColor Green

        try {
            $outputPath = Get-SafeOutputPath -SourcePath $fileInfo[0].Path -TargetBrand $targetBrand -IsMerge $true

            Write-Host ""
            Write-Host "  Writing merged output..." -NoNewline
            Write-AddressBook -Contacts $uniqueContacts -OutputPath $outputPath -TargetBrand $targetBrand
            Write-Host " Done" -ForegroundColor Green

            Write-Host "  Output: " -NoNewline -ForegroundColor White
            Write-Host $outputPath -ForegroundColor Cyan
        }
        catch {
            Write-Host " FAILED" -ForegroundColor Red
            Write-Host "  Error: $_" -ForegroundColor Red
        }
    }
    else {
        $fileCount = 1
        $totalFiles = $fileInfo.Count

        foreach ($info in $fileInfo) {
            if ($totalFiles -gt 1) {
                Write-Host ""
                Write-Host "[File $fileCount of $totalFiles]" -ForegroundColor Gray
            }

            $converted = Convert-AddressBook -SourcePath $info.Path -SourceBrand $info.Brand -TargetBrand $targetBrand

            if ($converted) {
                $allConvertedContacts += $converted
            }

            $fileCount++
        }
    }

    Show-ValidationReport

    Write-Log -Level 'INFO' -Function 'Main' -Message "Session completed"
    Write-FunctionExit -FunctionName 'Main'
    Write-Host ""
}

Main

#endregion
