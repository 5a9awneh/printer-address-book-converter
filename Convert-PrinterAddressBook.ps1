<#
.SYNOPSIS
    Printer Address Book Converter v2.0 - Universal printer contact list converter

.DESCRIPTION
    Converts printer address book files between Canon, Sharp, Xerox, and Develop/Konica formats.
    Auto-detects source format and handles validation, deduplication, and format conversion.
    Features intelligent fuzzy name matching, Outlook compatibility validation, and template-based conversion.

.PARAMETER SourcePath
    Path to source CSV file (for non-interactive mode)

.PARAMETER TargetPath
    Path to target CSV file - format will be auto-detected (for non-interactive mode)

.PARAMETER Mode
    Conversion mode: Single, Batch, or Merge (default: Single)

.PARAMETER NoInteractive
    Run in non-interactive mode without GUI prompts

.EXAMPLE
    .\Convert-PrinterAddressBook.ps1
    Interactive mode with menu navigation

.EXAMPLE
    .\Convert-PrinterAddressBook.ps1 -SourcePath "export.csv" -TargetPath "output.csv" -NoInteractive
    Non-interactive conversion to target format (detected from output.csv)

.NOTES
    Author: Faris Khasawneh
    Created: January 2026
    Version: 2.0
    Supports: Canon (iR-ADV, imageFORCE), Sharp MX/BP, Xerox AltaLink/VersaLink, Develop/Konica/Bizhub
    
    IMPORTANT: All CSV files (source and target template) must contain at least 1 valid contact
               with an email address for accurate format detection and conversion.

.CHANGELOG
    v2.0 - Phase 4: Intelligent deduplication + Outlook validation
           - Enhanced fuzzy name matching with edit distance algorithm
           - Detects abbreviations ("John Smith" vs "J. Smith") and typos
           - Outlook compatibility validation (email/name length, problematic characters)
           - Comprehensive test suite (brand conversions + edge cases)
           - Refactored output writing pipeline with quality validation
           - Unified success output with folder open prompt
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
    [string[]]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$TargetPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Single', 'Batch', 'Merge', 'Outlook')]
    [string]$Mode = 'Single',

    [Parameter(Mandatory = $false)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $false)]
    [switch]$NoInteractive
)

#region Configuration

$Script:BrandConfig = @{
    'Canon'   = @{
        NameField         = 'cn'
        EmailField        = 'mailaddress'
        Encoding          = 'UTF8'
        Delimiter         = ','
        HasComments       = $true
        HasQuotedHeaders  = $false  # Canon uses unquoted headers
        SignatureColumns  = @('objectclass', 'cn', 'mailaddress')
        # Output field mappings - defines what fields to populate and how
        OutputFields      = @{
            DisplayName = 'cn'
            Email       = 'mailaddress'
        }
        # Additional derived fields (computed from DisplayName)
        DerivedFields     = @{
            'cnread'  = { param($contact) $contact.DisplayName }
            'cnshort' = { param($contact) $contact.DisplayName.Substring(0, [Math]::Min(13, $contact.DisplayName.Length)) }
        }
        # Parsing rules
        FilterComments    = $true  # Filter lines starting with #
        HeaderPattern     = 'objectclass.*cn.*mailaddress'  # Pattern to find header line
        # Brand-specific header block for output
        BrandHeader       = "# Canon AddressBook CSV version: 0x0003`r`n# CharSet: UTF-8`r`n# dn: fixed`r`n# DB Version: 0x010b`r`n`r`n"
        TrailingDelimiter = $true  # Add delimiter after last column
    }
    'Sharp'   = @{
        NameField         = 'name'
        EmailField        = 'mail-address'
        Encoding          = 'UTF8'
        Delimiter         = ','
        HasComments       = $false
        HasQuotedHeaders  = $true   # Sharp uses quoted headers
        SignatureColumns  = @('address', 'name', 'mail-address', 'ftp-host')
        # Output field mappings - defines what fields to populate and how
        OutputFields      = @{
            DisplayName = 'name'
            Email       = 'mail-address'
        }
        # Additional derived fields (computed from DisplayName)
        DerivedFields     = @{
            'search-string' = { param($contact) $contact.DisplayName }
        }
        # Parsing rules
        FilterComments    = $false
        HeaderPattern     = $null  # Use last non-empty header line
        # Brand-specific header block for output
        BrandHeader       = ''  # No brand header
        TrailingDelimiter = $false
    }
    'Xerox'   = @{
        NameField         = 'DisplayName'
        NameFieldAlt      = @('FirstName', 'LastName')
        EmailField        = 'E-mailAddress'
        Encoding          = 'UTF8'
        Delimiter         = ','
        HasComments       = $false
        HasQuotedHeaders  = $false  # Xerox uses unquoted headers
        SignatureColumns  = @('XrxAddressBookId', 'DisplayName', 'E-mailAddress')
        # Output field mappings
        OutputFields      = @{
            DisplayName = 'DisplayName'
            FirstName   = 'FirstName'
            LastName    = 'LastName'
            Email       = 'E-mailAddress'
        }
        # Parsing rules
        FilterComments    = $false
        HeaderPattern     = $null  # Use last non-empty header line
        # Brand-specific header block for output
        BrandHeader       = ''  # No brand header
        TrailingDelimiter = $false
    }
    'Develop' = @{
        NameField         = 'Name'
        EmailField        = 'MailAddress'
        Encoding          = 'Unicode'
        Delimiter         = "`t"
        HasComments       = $false
        HasQuotedHeaders  = $false  # Develop uses unquoted headers
        SkipRows          = 2
        SignatureColumns  = @('AbbrNo', 'Name', 'MailAddress')
        # Output field mappings - defines what fields to populate and how
        OutputFields      = @{
            DisplayName = 'Name'
            Email       = 'MailAddress'
        }
        # Additional derived fields (computed from DisplayName)
        DerivedFields     = @{
            'Furigana'  = { param($contact) $contact.DisplayName }
            'SearchKey' = { param($contact) Get-SearchKey -Name $contact.DisplayName }
        }
        # Parsing rules
        FilterComments    = $false
        FilterAlternative = $true  # Skip "alternative" row
        HeaderPattern     = 'AbbrNo\tName\t'  # Pattern to find header line
        # Brand-specific header block for output (dynamic timestamp)
        BrandHeader       = { 
            param($timestamp)
            "@Ver406`t22C-6e`tIntegrate`tUTF-16LE`t${timestamp}`tabbr`t000ac209552c58a5b6222cf539ab712a255b`t`r`n#Abbreviate`t2000`t`r`n"
        }
        TrailingDelimiter = $true  # Add delimiter after last column
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
    
    .PARAMETER SourceMapping
        Mapping of logical fields (Email, Name) to CSV headers
    
    .OUTPUTS
        Normalized contact hashtable: @{Email, FirstName, LastName, DisplayName}
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Contact,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$SourceMapping
    )

    Write-FunctionEntry -FunctionName 'ConvertTo-NormalizedContact' -Parameters @{ Mapping = ($SourceMapping | Out-String) }

    try {
        # Extract email
        $emailField = $SourceMapping['Email']
        if (-not $emailField) { return $null }
        
        $email = $Contact.$emailField
        if ([string]::IsNullOrWhiteSpace($email)) {
            # Try alternatives if mapping had multiple? No, simpler to return null.
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
        
        if ($SourceMapping['IsSplitName']) {
            $firstName = $Contact.($SourceMapping['FirstName'])
            $lastName = $Contact.($SourceMapping['LastName'])
            
            if (-not [string]::IsNullOrWhiteSpace($firstName) -or -not [string]::IsNullOrWhiteSpace($lastName)) {
                $rawName = "$firstName $lastName".Trim()
                $displayName = (Get-Culture).TextInfo.ToTitleCase($rawName.ToLower())
                
                # Update first/last with normalized case?
                # Yes, assuming we want clean data 
                if ($firstName) { $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName.ToLower()) }
                if ($lastName) { $lastName = (Get-Culture).TextInfo.ToTitleCase($lastName.ToLower()) }
            }
        }
        elseif ($SourceMapping['DisplayName']) {
            $displayNameField = $SourceMapping['DisplayName']
            $displayName = $Contact.$displayNameField
            
            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = $displayName.Trim()
                # Normalize capitalization
                $displayName = (Get-Culture).TextInfo.ToTitleCase($displayName.ToLower())
                
                # Determine splitting strategy?
                # For now, always split for normalized data
                $nameParts = Split-FullName -FullName $displayName
                $firstName = $nameParts.FirstName
                $lastName = $nameParts.LastName
            }
        }
        
        # Fallback Name = Email User
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $nameParts = $email -split '@'
            $displayName = $nameParts[0]
            # Normalize capitalization
            $displayName = (Get-Culture).TextInfo.ToTitleCase($displayName.ToLower())
            $firstName = $displayName
        }
        
        # Create normalized contact
        $normalized = @{
            Email       = $email
            FirstName   = $firstName
            LastName    = $lastName
            DisplayName = $displayName
        }
        
        Write-Log -Level 'DEBUG' -Function 'ConvertTo-NormalizedContact' -Message "Normalized: $($normalized.DisplayName) <$($normalized.Email)>"
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
        
    .PARAMETER TemplateMapping
        Optional mapping from header analysis for template-based conversion
        
    .PARAMETER TemplateHeaders
        Optional list of all headers in the template (to ensure all fields exist)
    
    .OUTPUTS
        Hashtable with target brand field names and values
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$NormalizedContact,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Canon', 'Sharp', 'Xerox', 'Develop')]
        [string]$TargetBrand,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$TemplateMapping,
        
        [Parameter(Mandatory = $false)]
        [string[]]$TemplateHeaders
    )

    Write-FunctionEntry -FunctionName 'ConvertFrom-NormalizedContact' -Parameters @{ TargetBrand = $TargetBrand; UseTemplate = [bool]$TemplateMapping; Email = $NormalizedContact.Email }

    try {
        $targetContact = [ordered]@{}
        
        # ---------------------------------------------------------
        # MODE 1: TEMPLATE-BASED CONVERSION
        # ---------------------------------------------------------
        if ($TemplateMapping -and $TemplateHeaders) {
            # Initialize all fields from template with empty strings
            foreach ($header in $TemplateHeaders) {
                $targetContact[$header] = ""
            }
            
            # Fill mapped fields
            if ($TemplateMapping.ContainsKey('Email')) {
                $targetContact[$TemplateMapping['Email']] = $NormalizedContact.Email
            }
            
            # Handle Name (Split vs Single)
            if ($TemplateMapping['IsSplitName'] -eq $true) {
                $targetContact[$TemplateMapping['FirstName']] = $NormalizedContact.FirstName
                $targetContact[$TemplateMapping['LastName']] = $NormalizedContact.LastName
            }
            elseif ($TemplateMapping.ContainsKey('DisplayName')) {
                $targetContact[$TemplateMapping['DisplayName']] = $NormalizedContact.DisplayName
            }
            
            # Handle Search Key
            if ($TemplateMapping.ContainsKey('SearchKey')) {
                # Calculate search key based on FirstName or DisplayName
                $searchBase = if ($NormalizedContact.FirstName) { $NormalizedContact.FirstName } else { $NormalizedContact.DisplayName }
                $targetContact[$TemplateMapping['SearchKey']] = Get-SearchKey -Name $searchBase
            }
            
            Write-Log -Level 'DEBUG' -Function 'ConvertFrom-NormalizedContact' -Message "Mapped to Template format: $($NormalizedContact.DisplayName)"
            Write-FunctionExit -FunctionName 'ConvertFrom-NormalizedContact' -Result "Template"
            
            return $targetContact
        }
        
        # ---------------------------------------------------------
        # MODE 2: BRAND-CONFIG CONVERSION (Legacy/Fallback)
        # ---------------------------------------------------------
        if (-not $TargetBrand) {
            throw "TargetBrand is required if no TemplateMapping is provided."
        }
        
        $config = $Script:BrandConfig[$TargetBrand]
        $targetContact = @{}
        
        # Map fields based on OutputFields configuration (brand-agnostic)
        foreach ($normalizedField in $config.OutputFields.Keys) {
            $targetFieldName = $config.OutputFields[$normalizedField]
            
            # Map normalized field to target value
            switch ($normalizedField) {
                'DisplayName' { $targetContact[$targetFieldName] = $NormalizedContact.DisplayName }
                'FirstName' { $targetContact[$targetFieldName] = $NormalizedContact.FirstName }
                'LastName' { $targetContact[$targetFieldName] = $NormalizedContact.LastName }
                'Email' { $targetContact[$targetFieldName] = $NormalizedContact.Email }
            }
        }
        
        # Add brand-specific derived fields if configured
        if ($config.DerivedFields) {
            foreach ($fieldName in $config.DerivedFields.Keys) {
                $scriptBlock = $config.DerivedFields[$fieldName]
                $targetContact[$fieldName] = & $scriptBlock $NormalizedContact
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

function Get-CsvDetails {
    <#
    .SYNOPSIS
        Detects CSV encoding and delimiter from a file.
    .OUTPUTS
        Hashtable with Encoding (string) and Delimiter (char)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    try {
        # 1. Detect Encoding (Basic heuristic)
        $bytes = Get-Content -Path $FilePath -Encoding Byte -TotalCount 2 -ErrorAction SilentlyContinue
        $encodingName = 'UTF8'
        if ($bytes -and $bytes.Count -ge 2) {
            if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
                $encodingName = 'Unicode'
            }
        }

        # 2. Detect Delimiter
        # Read the first few lines as text to guess delimiter
        $sampleLines = Get-Content -Path $FilePath -Encoding $encodingName -TotalCount 5
        $headerLine = $sampleLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1

        $delimiter = ',' # Default
        if ($headerLine) {
            $commaCount = ($headerLine.ToCharArray() | Where-Object { $_ -eq ',' }).Count
            $tabCount = ($headerLine.ToCharArray() | Where-Object { $_ -eq "`t" }).Count
            $semiCount = ($headerLine.ToCharArray() | Where-Object { $_ -eq ';' }).Count

            # Heuristic: The delimiter must appear at least once (unless single column, which is edge case)
            if ($tabCount -gt $commaCount -and $tabCount -gt $semiCount -and $tabCount -gt 0) {
                $delimiter = "`t"
            }
            elseif ($semiCount -gt $commaCount -and $semiCount -gt 0) {
                $delimiter = ';'
            }
            # Else defaults to comma
        }

        return @{
            Encoding  = $encodingName
            Delimiter = $delimiter
        }
    }
    catch {
        Write-Log -Level 'WARN' -Function 'Get-CsvDetails' -Message "Failed to detect CSV details: $_"
        return @{ Encoding = 'UTF8'; Delimiter = ',' }
    }
}

function Get-TemplateMapping {
    <#
    .SYNOPSIS
        Heuristically maps normalized fields (Email, Name) to CSV headers.
    .OUTPUTS
        Hashtable mapping NormalizedField -> CsvHeaderName
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Headers
    )

    $mapping = @{}
    
    # Helper to find best match
    function Find-Header {
        param($Patterns)
        foreach ($pattern in $Patterns) {
            $match = $Headers | Where-Object { $_ -match $pattern } | Select-Object -First 1
            if ($match) { return $match }
        }
        return $null
    }

    # 1. Email (Priority)
    $emailHeader = Find-Header -Patterns @('(?i)^e?[-_]?mail', '(?i)addr(ess)?', '(?i)mail')
    if ($emailHeader) { $mapping['Email'] = $emailHeader }

    # 2. Name (Full vs First/Last)
    # Check for Split names first
    $firstHeader = Find-Header -Patterns @('(?i)^first[-_]?name', '(?i)^given[-_]?name', '(?i)^f[-_]name')
    $lastHeader = Find-Header -Patterns @('(?i)^last[-_]?name', '(?i)^sur[-_]?name', '(?i)^family[-_]?name', '(?i)^l[-_]name')

    if ($firstHeader -and $lastHeader) {
        $mapping['FirstName'] = $firstHeader
        $mapping['LastName'] = $lastHeader
        $mapping['IsSplitName'] = $true
    }
    else {
        # Fallback to display name (include 'cn' for LDAP/Canon styles)
        $nameHeader = Find-Header -Patterns @('(?i)^name', '(?i)^display[-_]?name', '(?i)^common[-_]?name', '(?i)^cn$', '(?i)^user[-_]?name', '(?i)^contact')
        if ($nameHeader) { $mapping['DisplayName'] = $nameHeader }
    }

    # 3. Search Key / Index / Yomi
    $searchHeader = Find-Header -Patterns @('(?i)search', '(?i)index', '(?i)yomi', '(?i)furi', '(?i)abc', '(?i)key', '(?i)read')
    if ($searchHeader) { $mapping['SearchKey'] = $searchHeader }

    return $mapping
}

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

        # Find first and last contact row indices (rows with email addresses)
        # Use email pattern instead of just @ to avoid false positives (e.g., "username/@encodingMethod")
        $firstContactIndex = -1
        $lastContactIndex = -1
        $emailPattern = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

        for ($i = 0; $i -lt $allLines.Count; $i++) {
            if ($allLines[$i] -match $emailPattern) {
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

        # Detect if headers are quoted by checking first non-comment header line
        $hasQuotedHeaders = $false
        $firstHeaderLine = $headers | Where-Object { -not ($_ -match '^#') -and -not ([string]::IsNullOrWhiteSpace($_)) } | Select-Object -First 1
        if ($firstHeaderLine) {
            # Check if line starts with quoted field: "field" or starts with tab then quote
            $hasQuotedHeaders = $firstHeaderLine -match '^"[^"]+"[,\t]' -or $firstHeaderLine -match '^\t?"[^"]+"'
        }

        $result = @{
            Headers          = $headers
            Contacts         = $contacts
            Footers          = $footers
            HasQuotedHeaders = $hasQuotedHeaders
        }
        
        Write-Log -Level 'DEBUG' -Function 'Get-CsvStructure' -Message "Quote detection: HasQuotedHeaders=$hasQuotedHeaders"
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

    try {
        # Use robust heuristic detection
        $details = Get-CsvDetails -FilePath $FilePath
        $structure = Get-CsvStructure -FilePath $FilePath -Encoding $details.Encoding

        if ($structure.Contacts.Count -eq 0 -and $structure.Headers.Count -eq 0) {
            return @{ Brand = 'Unknown'; Confidence = 0 }
        }

        # Find header line (skipping specific garbage)
        $headerLine = $structure.Headers | 
        Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_) -and 
            -not ($_ -match '^#') -and 
            -not ($_ -match '^"alternative"') -and
            -not ($_ -match '^@Ver')
        } | Select-Object -Last 1
            
        if (-not $headerLine) { 
            # Fallback: if no headers found before contacts, maybe the first line was it but it was complex?
            # Or file starts directly with contacts? unlikely.
            return @{ Brand = 'Unknown'; Confidence = 0 } 
        }

        # Parse columns based on detected delimiter
        $columns = if ($structure.HasQuotedHeaders) {
            ($headerLine -split $details.Delimiter) | ForEach-Object { $_.Trim().Trim('"') }
        }
        else {
            ($headerLine -split $details.Delimiter) | ForEach-Object { $_.Trim() }
        }

        foreach ($brand in $Script:BrandConfig.Keys) {
            $config = $Script:BrandConfig[$brand]
            $matchCount = 0

            foreach ($sigCol in $config.SignatureColumns) {
                if ($columns -contains $sigCol) {
                    $matchCount++
                }
            }
            
            # Use strict matching for identification
            if ($config.SignatureColumns.Count -gt 0) {
                $confidence = [math]::Round(($matchCount / $config.SignatureColumns.Count) * 100)
            }
            else {
                $confidence = 0
            }

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
        Write-Log "ERROR" "Detection failed for ${FilePath}: $_"
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

        # Parse contact lines based on brand format (using config metadata)
        $config = $Script:BrandConfig[$Brand]
        
        # Filter contact lines if needed
        $contactLines = $structure.Contacts
        if ($config.FilterComments) {
            $contactLines = $contactLines | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }
        }
        if ($config.FilterAlternative) {
            $contactLines = $contactLines | Where-Object { $_ -notmatch '^"alternative"' }
        }
        
        if ($contactLines.Count -eq 0) {
            Write-Log "WARN" "No data lines in $Brand file after filtering"
            return @()
        }
        
        # Find the column header line from Headers section
        $headerLine = $structure.Headers | 
        Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_) -and 
            -not ($_ -match '^#') -and 
            -not ($_ -match '^"alternative"') -and
            -not ($_ -match '^@Ver')
        } | Select-Object -Last 1
        
        if ([string]::IsNullOrWhiteSpace($headerLine)) {
            Write-Log "ERROR" "No column header found in $Brand file"
            return @()
        }
        
        # Create temp file with header + contact lines
        $tempFile = [System.IO.Path]::GetTempFileName()
        @($headerLine) + $contactLines | Out-File -FilePath $tempFile -Encoding $config.Encoding
        $data = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
        Remove-Item -Path $tempFile -Force

        # Extract email and name from each row (brand-agnostic using config metadata)
        foreach ($row in $data) {
            $email = $row.($config.EmailField)

            if ([string]::IsNullOrWhiteSpace($email)) {
                continue
            }

            # Handle name fields using config metadata
            if ($config.NameFieldAlt) {
                # Brand has separate FirstName/LastName fields (e.g., Xerox)
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
                # Single name field
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

function Show-OutputSuccess {
    <#
    .SYNOPSIS
        Displays a unified success message with file details and offers to open the folder.
    
    .DESCRIPTION
        Shows consistent success output across all modes (Convert, Merge, Outlook).
        Displays file path, size, contact count, and prompts to open in Explorer.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [int]$ContactCount = 0,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoPrompt,
        
        [Parameter(Mandatory = $false)]
        [switch]$Silent
    )
    
    # Silent mode: don't show any output (used in batch processing)
    if ($Silent) {
        return
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SUCCESS! File created:" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Location: " -NoNewline -ForegroundColor White
    Write-Host $OutputPath -ForegroundColor Cyan
    Write-Host ""
    
    # Verify file exists and show details
    if (Test-Path $OutputPath) {
        $fileInfo = Get-Item $OutputPath
        Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor Gray
        if ($ContactCount -gt 0) {
            Write-Host "Contacts written: $ContactCount" -ForegroundColor Gray
        }
        Write-Host ""
        
        # Offer to open folder (unless suppressed)
        if (-not $NoPrompt) {
            $openFolder = Read-Host "Open folder in Explorer? (Y/N)"
            if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
                Start-Process "explorer.exe" -ArgumentList "/select,`"$OutputPath`""
            }
        }
    }
    else {
        Write-Host "WARNING: File not found at expected location!" -ForegroundColor Yellow
    }
}

function Write-AddressBook {
    <#
    .SYNOPSIS
        Writes normalized contacts to target brand CSV format.
    
    .DESCRIPTION
        Accepts normalized contact objects and converts them to target brand format.
        Uses ConvertFrom-NormalizedContact for field mapping.
        Supports both Legacy BrandConfig mode and new Template-based mode.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$NormalizedContacts,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetBrand,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$TemplateStructure,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$TemplateDetails,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$TemplateMapping,

        [Parameter(Mandatory = $false)]
        [string[]]$TemplateColumns
    )

    Write-FunctionEntry -FunctionName 'Write-AddressBook' -Parameters @{ OutputPath = $OutputPath; TargetBrand = $TargetBrand; ContactCount = $NormalizedContacts.Count; UseTemplate = [bool]$TemplateStructure }

    try {
        # ---------------------------------------------------------
        # MODE 1: TEMPLATE-BASED WRITING
        # ---------------------------------------------------------
        if ($TemplateStructure -and $TemplateColumns) {
            Write-Log -Level 'INFO' -Function 'Write-AddressBook' -Message "Writing using Template mode"
            
            # Determine Encoding
            $encoding = [System.Text.Encoding]::UTF8
            if ($TemplateDetails.Encoding -eq 'Unicode') { 
                $encoding = [System.Text.Encoding]::Unicode 
            }
            elseif ($TemplateDetails.Encoding -eq 'BigEndianUnicode') {
                $encoding = [System.Text.Encoding]::BigEndianUnicode
            }

            # 1. Write Header Rows (Preserved from template)
            if ($TemplateStructure.Headers.Count -gt 0) {
                [System.IO.File]::WriteAllLines($OutputPath, $TemplateStructure.Headers, $encoding)
            }
            else {
                # Fallback: Write columns if no existing headers detected (unlikely)
                $headerLine = ($TemplateColumns | ForEach-Object { "`"$_`"" }) -join $TemplateDetails.Delimiter
                [System.IO.File]::WriteAllLines($OutputPath, @($headerLine), $encoding)
            }

            # 2. Process and Write Contacts
            foreach ($normalized in $NormalizedContacts) {
                $targetFields = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TemplateMapping $TemplateMapping -TemplateHeaders $TemplateColumns
                
                # Build CSV Line
                $values = foreach ($col in $TemplateColumns) {
                    $val = $targetFields[$col]
                    if ($null -eq $val) { $val = '' }
                    
                    # Handle Quotes: 
                    # If the template used quoted headers, we should probably quote values too.
                    # Or we just always quote values to be safe (safest approach for CSV).
                    # Escape existing quotes
                    $val = $val -replace '"', '""'
                    "`"$val`""
                }
                
                $line = $values -join $TemplateDetails.Delimiter
                [System.IO.File]::AppendAllText($OutputPath, $line + "`r`n", $encoding)
            }

            # 3. Write Footer Rows
            if ($TemplateStructure.Footers.Count -gt 0) {
                $footerContent = ($TemplateStructure.Footers -join "`r`n") + "`r`n"
                [System.IO.File]::AppendAllText($OutputPath, $footerContent, $encoding)
            }

            Write-Log -Level 'INFO' -Function 'Write-AddressBook' -Message "Wrote $($NormalizedContacts.Count) contacts (Template Mode)"
            Write-FunctionExit -FunctionName 'Write-AddressBook' -Result "$($NormalizedContacts.Count) contacts written"
            return
        }

        # ---------------------------------------------------------
        # MODE 2: LEGACY BRAND-CONFIG WRITING
        # ---------------------------------------------------------
        if (-not $TargetBrand) { throw "TargetBrand required if no template provided" }

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

        # Get target brand's quote style from configuration
        $useQuotedHeaders = $Script:BrandConfig[$TargetBrand].HasQuotedHeaders
        Write-Log -Level 'DEBUG' -Function 'Write-AddressBook' -Message "Target brand '$TargetBrand' uses quoted headers: $useQuotedHeaders"

        # Use Export-Csv for Sharp (quoted headers), manual writing for all others (unquoted headers)
        if ($useQuotedHeaders) {
            # Sharp: Use PowerShell Export-Csv (adds quotes automatically)
            $output | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Log -Level 'DEBUG' -Function 'Write-AddressBook' -Message "Used Export-Csv for quoted format"
        }
        else {
            # Canon, Xerox, Develop: Manual writing with unquoted headers
            Write-Log -Level 'DEBUG' -Function 'Write-AddressBook' -Message "Using manual writing for unquoted format"
            
            # Get encoding and delimiter from config
            $config = $Script:BrandConfig[$TargetBrand]
            $encoding = [System.Text.Encoding]::($config.Encoding)
            $delimiter = $config.Delimiter
            
            # Get brand-specific header (evaluate scriptblock if needed)
            $brandHeader = if ($config.BrandHeader -is [scriptblock]) {
                $timestamp = Get-Date -Format 'yyyy.M.d HH:mm:ss'
                & $config.BrandHeader $timestamp
            }
            else {
                $config.BrandHeader
            }

            # Get column names from first object
            if ($output.Count -gt 0) {
                $columns = $output[0].PSObject.Properties.Name
                
                # Build unquoted column header
                $columnHeader = $columns -join $delimiter
                if ($config.TrailingDelimiter) {
                    $columnHeader += $delimiter
                }
                
                # Write brand header + column header
                $headerBlock = $brandHeader + $columnHeader + "`r`n"
                [System.IO.File]::WriteAllText($OutputPath, $headerBlock, $encoding)
                
                # Write data rows (quoted values)
                foreach ($obj in $output) {
                    $values = foreach ($col in $columns) {
                        $val = $obj.$col
                        if ($null -eq $val) { $val = '' }
                        "`"$val`""
                    }
                    $line = $values -join $delimiter
                    [System.IO.File]::AppendAllText($OutputPath, $line + "`r`n", $encoding)
                }
            }
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
    <#
    .SYNOPSIS
        Removes duplicate contacts based on email address.
    
    .DESCRIPTION
        Simple email-based deduplication (case-insensitive).
    
    .PARAMETER Contacts
        Array of contact objects to deduplicate
    
    .OUTPUTS
        Array of unique contacts
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$Contacts
    )

    Write-FunctionEntry -FunctionName 'Remove-Duplicates' -Parameters @{ ContactCount = $Contacts.Count }

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
        Write-Log -Level 'INFO' -Function 'Remove-Duplicates' -Message "Removed $duplicates duplicates"
    }

    Write-FunctionExit -FunctionName 'Remove-Duplicates' -Result "$($unique.Count) unique contacts"
    return $unique
}

function Get-NameSimilarity {
    <#
    .SYNOPSIS
        Calculates similarity between two names using edit distance.
    
    .DESCRIPTION
        Returns a similarity score between 0.0 (completely different) and 1.0 (identical).
        Uses normalized edit distance for fuzzy matching.
    
    .PARAMETER Name1
        First name to compare
    
    .PARAMETER Name2
        Second name to compare
    
    .OUTPUTS
        Float between 0.0 and 1.0 representing similarity
    #>
    param(
        [string]$Name1,
        [string]$Name2
    )

    if ([string]::IsNullOrWhiteSpace($Name1) -or [string]::IsNullOrWhiteSpace($Name2)) {
        return 0.0
    }

    $name1Lower = $Name1.ToLower().Trim()
    $name2Lower = $Name2.ToLower().Trim()

    # Exact match
    if ($name1Lower -eq $name2Lower) {
        return 1.0
    }

    # Check for abbreviation matches (e.g., "John Smith" vs "J. Smith")
    $name1Parts = $name1Lower -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $name2Parts = $name2Lower -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
    if ($name1Parts.Count -gt 0 -and $name2Parts.Count -gt 0) {
        $last1 = $name1Parts | Select-Object -Last 1
        $last2 = $name2Parts | Select-Object -Last 1

        if ($last1 -eq $last2) {
            # Last names match - check if first name is abbreviated
            $first1 = $name1Parts | Select-Object -First 1
            $first2 = $name2Parts | Select-Object -First 1
            
            if ($first1 -and $first2 -and $first1.Length -gt 0 -and $first2.Length -gt 0) {
                if ($first1.Substring(0, 1) -eq $first2.Substring(0, 1)) {
                    return 0.85
                }
            }
        }
    }

    # Calculate edit distance
    $distance = Get-EditDistance -String1 $name1Lower -String2 $name2Lower
    $maxLength = [Math]::Max($name1Lower.Length, $name2Lower.Length)
    
    if ($maxLength -eq 0) {
        return 1.0
    }

    # Normalize to 0.0 - 1.0 range (1.0 = identical, 0.0 = completely different)
    $similarity = 1.0 - ($distance / $maxLength)
    return $similarity
}

function Get-EditDistance {
    <#
    .SYNOPSIS
        Calculates edit distance between two strings.
    
    .DESCRIPTION
        Returns the minimum number of single-character edits (insertions, deletions, substitutions)
        required to change one string into another.
    
    .PARAMETER String1
        First string
    
    .PARAMETER String2
        Second string
    
    .OUTPUTS
        Integer representing edit distance
    #>
    param(
        [string]$String1,
        [string]$String2
    )

    # Normalize nulls to empty strings to avoid index issues
    if ($null -eq $String1) { $String1 = '' }
    if ($null -eq $String2) { $String2 = '' }

    Write-Host "Debug: EditDistance s1='$String1' (type=$($String1.GetType().FullName)) s2='$String2' (type=$($String2.GetType().FullName))"

    $len1 = $String1.Length
    $len2 = $String2.Length

    # Use a real 2D array to avoid jagged array indexing issues
    $matrix = [int[, ]]::new($len1 + 1, $len2 + 1)
    Write-Host "Debug: matrix type=$($matrix.GetType().FullName) dims=($($matrix.GetLength(0)))x($($matrix.GetLength(1)))"

    # Initialize first column and row
    for ($i = 0; $i -le $len1; $i++) {
        $matrix[$i, 0] = $i
    }
    for ($j = 0; $j -le $len2; $j++) {
        $matrix[0, $j] = $j
    }

    # Fill matrix
    for ($i = 1; $i -le $len1; $i++) {
        for ($j = 1; $j -le $len2; $j++) {
            try {
                $cost = if ($String1[$i - 1] -eq $String2[$j - 1]) { 0 } else { 1 }
                
                $deletion = $matrix[$i - 1, $j] + 1
                $insertion = $matrix[$i, $j - 1] + 1
                $substitution = $matrix[$i - 1, $j - 1] + $cost
                
                $matrix[$i, $j] = [Math]::Min([Math]::Min($deletion, $insertion), $substitution)
            }
            catch {
                Write-Host "Debug: EditDistance inner error i=$i (type=$($i.GetType().FullName)) j=$j (type=$($j.GetType().FullName))"
                Write-Host "Debug: String1[$i-1]=$($String1[$i - 1]) (type=$($String1[$i - 1].GetType().FullName))"
                Write-Host "Debug: String2[$j-1]=$($String2[$j - 1]) (type=$($String2[$j - 1].GetType().FullName))"
                Write-Host "Debug: matrix element types del=$([string]($matrix[$i - 1,$j]).GetType().FullName) ins=$([string]($matrix[$i,$j - 1]).GetType().FullName) sub=$([string]($matrix[$i - 1,$j - 1]).GetType().FullName)"
                throw
            }
        }
    }

    return $matrix[$len1, $len2]
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
    
    # Skip strict validation for generic "Converted" output
    if ($TargetBrand -eq 'Converted' -or -not ($Script:BrandConfig.ContainsKey($TargetBrand))) {
        Write-Log -Level 'INFO' -Function 'Validate-OutputFile' -Message "Skipping strict validation for generic/unknown brand: $TargetBrand"
        return $result
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

        # Get expected encoding for brand from config
        $config = $Script:BrandConfig[$TargetBrand]
        $expectedEncoding = $config.Encoding

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

        # Parse contacts to check field structure (brand-agnostic using config metadata)
        try {
            # Filter contact lines if needed
            $contactLines = $structure.Contacts
            if ($config.FilterComments) {
                $contactLines = $contactLines | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }
            }
            
            if ($contactLines.Count -gt 0) {
                $tempFile = [System.IO.Path]::GetTempFileName()
                $contactLines | Out-File -FilePath $tempFile -Encoding $expectedEncoding
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

function Test-OutlookCompatibility {
    <#
    .SYNOPSIS
        Validates contacts for Microsoft Outlook import compatibility.
    
    .DESCRIPTION
        Checks for common issues that prevent successful Outlook import:
        - Email address length limits (max 254 characters)
        - Name field length limits (max 256 characters)
        - Problematic characters in email or name fields
        - Required field validation
        - Special character encoding issues
    
    .PARAMETER Contacts
        Array of normalized contact objects to validate
    
    .OUTPUTS
        Hashtable with validation results: @{IsCompatible, Errors[], Warnings[], ContactCount}
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$Contacts
    )

    Write-FunctionEntry -FunctionName 'Test-OutlookCompatibility' -Parameters @{ ContactCount = $Contacts.Count }

    $result = @{
        IsCompatible = $true
        Errors       = @()
        Warnings     = @()
        ContactCount = $Contacts.Count
        IssueCount   = 0
    }

    # Outlook limits
    $maxEmailLength = 254
    $maxNameLength = 256
    $problematicChars = @('[', ']', '<', '>', ';', ':', ',', '"', '\', '/')

    foreach ($contact in $Contacts) {
        $contactIssues = @()
        
        # Check email length
        if ($contact.Email.Length -gt $maxEmailLength) {
            $contactIssues += "Email exceeds $maxEmailLength characters ($($contact.Email.Length))"
            $result.IsCompatible = $false
        }

        # Check email for problematic characters
        foreach ($char in $problematicChars) {
            if ($contact.Email.Contains($char)) {
                $contactIssues += "Email contains problematic character: '$char'"
                $result.Warnings += "Contact '$($contact.DisplayName)': Email contains '$char'"
            }
        }

        # Check name length
        if ($contact.DisplayName.Length -gt $maxNameLength) {
            $contactIssues += "Name exceeds $maxNameLength characters ($($contact.DisplayName.Length))"
            $result.IsCompatible = $false
        }

        # Check for control characters
        if ($contact.Email -match '[\x00-\x1F\x7F]') {
            $contactIssues += "Email contains control characters"
            $result.IsCompatible = $false
        }

        if ($contact.DisplayName -match '[\x00-\x1F\x7F]') {
            $contactIssues += "Name contains control characters"
            $result.Warnings += "Contact '$($contact.DisplayName)': Name has control characters"
        }

        # Check for leading/trailing spaces
        if ($contact.Email -ne $contact.Email.Trim()) {
            $contactIssues += "Email has leading or trailing spaces"
            $result.Warnings += "Contact '$($contact.DisplayName)': Email needs trimming"
        }

        # Log issues for this contact
        if ($contactIssues.Count -gt 0) {
            $result.IssueCount++
            $issueMsg = "Contact '$($contact.DisplayName)' <$($contact.Email)>: $($contactIssues -join ', ')"
            $result.Errors += $issueMsg
            Write-Log -Level 'WARN' -Function 'Test-OutlookCompatibility' -Message $issueMsg
        }
    }

    if ($result.IsCompatible) {
        Write-Log -Level 'INFO' -Function 'Test-OutlookCompatibility' -Message "All $($Contacts.Count) contacts are Outlook-compatible"
    }
    else {
        Write-Log -Level 'WARN' -Function 'Test-OutlookCompatibility' -Message "$($result.IssueCount) contacts have compatibility issues"
    }

    Write-FunctionExit -FunctionName 'Test-OutlookCompatibility' -Result "IsCompatible=$($result.IsCompatible), Issues=$($result.IssueCount)"
    return $result
}

function Parse-OutlookContacts {
    <#
    .SYNOPSIS
        Parses text in Microsoft Outlook "Check Names" format.
    .DESCRIPTION
        Format: "LAST, First <email@domain.com>; ..."
        Returns custom objects ready for normalization.
    #>
    param([string]$InputText)
    
    $contacts = @()
    # Replace newlines with spaces to handle wrapped text
    $cleanText = $InputText -replace "`r`n", " " -replace "`n", " "
    
    $cleanText -split ';' | ForEach-Object {
        if ($_ -match '(.+?)<(.+?)>') {
            $nameRaw = $matches[1].Trim()
            $email = $matches[2].Trim().ToLower()

            # Convert "LASTNAME, Firstname" to "Firstname LASTNAME"
            if ($nameRaw -match '^(.+?),\s*(.+)$') {
                $nameRaw = "$($matches[2].Trim()) $($matches[1].Trim())"
            }
            
            if (Test-Email -Email $email) {
                $contacts += [PSCustomObject]@{
                    Name  = $nameRaw
                    Email = $email
                }
            }
        }
    }
    return $contacts
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
        Write-Host "  PRINTER ADDRESS BOOK CONVERTER v2.0" -ForegroundColor Cyan
        Write-Host "  Canon | Sharp | Xerox | Develop/Konica" -ForegroundColor Cyan
        Write-Host "  by Faris Khasawneh - January 2026" -ForegroundColor Gray
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""

        if ($Description) {
            # Split description to handle warning text with yellow color
            $lines = $Description -split "`n"
            foreach ($line in $lines) {
                if ($line -match '^⚠️') {
                    Write-Host $line -ForegroundColor Yellow
                }
                else {
                    Write-Host $line -ForegroundColor Gray
                }
            }
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

        if ($IsMerge) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
            $fileName = "Merged_converted_${timestamp}.csv"
        }
        else {
            $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
            $sourceFileName = $sourceFileName -replace '[\\/:*?"<>|]', '_'
            $fileName = "${sourceFileName}_converted.csv"
        }

        # Check if file exists and append number if needed
        $fullPath = Join-Path -Path $sourceDir -ChildPath $fileName
        if (Test-Path $fullPath) {
            $counter = 2
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $extension = [System.IO.Path]::GetExtension($fileName)
            
            do {
                $fileName = "${baseName}_${counter}${extension}"
                $fullPath = Join-Path -Path $sourceDir -ChildPath $fileName
                $counter++
            } while (Test-Path $fullPath)
        }

        return $fullPath
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
        Brand of target format (Optional if TemplatePath provided)

    .PARAMETER TargetTemplatePath
        Path to an existing CSV implementation to use as a template (Required)
    
    .OUTPUTS
        Array of processed normalized contacts
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceBrand,
        
        [Parameter(Mandatory = $false)]
        [string]$TargetBrand,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetTemplatePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$SuppressPrompt
    )

    Write-FunctionEntry -FunctionName 'Convert-AddressBook' -Parameters @{ SourcePath = $SourcePath; SourceBrand = $SourceBrand; TargetBrand = $TargetBrand; Template = $TargetTemplatePath }

    Write-Host ""
    Write-Host "Processing: " -NoNewline
    Write-Host (Split-Path -Leaf $SourcePath) -ForegroundColor Cyan
    Write-Host "  Source: " -NoNewline
    Write-Host $SourceBrand -ForegroundColor Green
    
    if ($TargetTemplatePath) {
        Write-Host "  Template: " -NoNewline
        Write-Host (Split-Path -Leaf $TargetTemplatePath) -ForegroundColor Green
    }
    else {
        Write-Host "  Target: " -NoNewline
        Write-Host $TargetBrand -ForegroundColor Green
    }

    # Step 1: Parse source CSV and get raw contact rows
    Write-Host "  Reading..." -NoNewline
    try {
        # 1. Detect Encoding and Delimiter
        $sourceDetails = Get-CsvDetails -FilePath $SourcePath
        
        # 2. Get Structure (Headers/Contacts)
        $sourceStructure = Get-CsvStructure -FilePath $SourcePath -Encoding $sourceDetails.Encoding
        
        if ($sourceStructure.Contacts.Count -eq 0) {
            Write-Host " No contacts found" -ForegroundColor Yellow
            return $null
        }
        
        # 3. Find Header Line (Look for 'Unknown' stuff like @Ver or "alternative")
        $headerLine = $sourceStructure.Headers | 
        Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_) -and 
            -not ($_ -match '^#') -and 
            -not ($_ -match '^"alternative"') -and
            -not ($_ -match '^@Ver')
        } | Select-Object -Last 1

        if (-not $headerLine) { throw "Missing header row" }
        
        # 4. Extract Columns
        $sourceCols = if ($sourceStructure.HasQuotedHeaders) {
            # Basic quoted CSV split - split, trim whitespace, then trim quotes
            ($headerLine -split $sourceDetails.Delimiter) | ForEach-Object { $_.Trim().Trim('"') }
        }
        else {
            ($headerLine -split $sourceDetails.Delimiter) | ForEach-Object { $_.Trim() }
        }
        
        # Filter out empty columns
        $sourceCols = $sourceCols | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        
        # 5. Get Mapping (Heuristic)
        $sourceMapping = Get-TemplateMapping -Headers $sourceCols
        if (-not $sourceMapping['Email']) {
            # Fallback to BrandConfig if available (to keep stability for weird files)
            if ($Script:BrandConfig.ContainsKey($SourceBrand)) {
                $config = $Script:BrandConfig[$SourceBrand]
                $sourceMapping['Email'] = $config.EmailField
                $sourceMapping['DisplayName'] = $config.NameField
            }
            else {
                throw "Could not identify Email column"
            }
        }
        
        # 6. Parse Rows
        $tempFile = [System.IO.Path]::GetTempFileName()
        # Combine valid header + valid contacts
        @($headerLine) + $sourceStructure.Contacts | Out-File -FilePath $tempFile -Encoding $sourceDetails.Encoding
        $sourceRows = Import-Csv -Path $tempFile -Delimiter $sourceDetails.Delimiter
        Remove-Item -Path $tempFile -Force
        
        Write-Host " $($sourceRows.Count) contacts" -ForegroundColor Green
        Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Read $($sourceRows.Count) contact rows"
    }
    catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        Write-Log -Level 'ERROR' -Function 'Convert-AddressBook' -Message "Read failed" -ErrorRecord $_
        return $null
    }

    # Step 2: Normalize contacts
    Write-Host "  Normalizing..." -NoNewline
    $normalizedContacts = @()
    foreach ($row in $sourceRows) {
        $normalized = ConvertTo-NormalizedContact -Contact $row -SourceMapping $sourceMapping
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

    # Step 3: Validate
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

    # Step 5: Prepare Output / Template Analysis
    try {
        $templateStructure = $null
        $templateDetails = $null
        $templateMapping = $null
        $templateColumns = $null

        if ($TargetTemplatePath) {
            Write-Host "  Analyzing Template..." -NoNewline
            
            # 5a. Analyze Template File
            $templateDetails = Get-CsvDetails -FilePath $TargetTemplatePath
            $templateStructure = Get-CsvStructure -FilePath $TargetTemplatePath -Encoding $templateDetails.Encoding
            
            # 5b. Extract Columns
            # Find the header line (avoiding comments and known metadata lines like "alternative")
            $colHeaderLine = $templateStructure.Headers | 
            Where-Object { 
                -not [string]::IsNullOrWhiteSpace($_) -and 
                -not ($_ -match '^#') -and 
                -not ($_ -match '^"alternative"') -and
                -not ($_ -match '^@Ver')
            } | Select-Object -Last 1
            
            if (-not $colHeaderLine) {
                throw "Template file has no recognized headers"
            }
            
            # Parse columns (Handle quotes if needed)
            if ($templateStructure.HasQuotedHeaders) {
                # Basic CSV parsing of the header line
                $templateColumns = ($colHeaderLine -split $templateDetails.Delimiter).TrimStart('"').TrimEnd('"')
            }
            else {
                $templateColumns = $colHeaderLine -split $templateDetails.Delimiter
            }
            
            # Remove empty columns (fixes issue with trailing delimiters)
            $templateColumns = $templateColumns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            # 5c. Generate Mapping
            $templateMapping = Get-TemplateMapping -Headers $templateColumns
            
            Write-Host " OK (Heuristic)" -ForegroundColor Green
            Write-Host "    Found Columns: " -NoNewline
            Write-Host "$($templateColumns.Count)" -ForegroundColor Cyan
            Write-Host "    Mapped Email: " -NoNewline
            if ($templateMapping.Email) { Write-Host $templateMapping.Email -ForegroundColor Cyan } else { Write-Host "Not Found" -ForegroundColor Red }
             
            # Use original filename logic but targeted to output
            $outputPath = Get-SafeOutputPath -SourcePath $SourcePath -TargetBrand "Converted"
            # Or overwrite? User said "remove data rows from target csv...". 
            # Usually safer to write to new file.
        }
        else {
            $outputPath = Get-SafeOutputPath -SourcePath $SourcePath -TargetBrand $TargetBrand
        }
        
        Write-Log -Level 'INFO' -Function 'Convert-AddressBook' -Message "Output path: $outputPath"
    }
    catch {
        Write-Host "  Analysis/Path error: $_" -ForegroundColor Red
        Write-Log -Level 'ERROR' -Function 'Convert-AddressBook' -Message "Output prep error" -ErrorRecord $_
        return $null
    }

    # Step 6: Write to target format
    Write-Host "  Writing..." -NoNewline
    try {
        if ($TargetTemplatePath) {
            Write-AddressBook -NormalizedContacts $uniqueContacts -OutputPath $outputPath `
                -TemplateStructure $templateStructure `
                -TemplateDetails $templateDetails `
                -TemplateMapping $templateMapping `
                -TemplateColumns $templateColumns
        }
        else {
            Write-AddressBook -NormalizedContacts $uniqueContacts -OutputPath $outputPath -TargetBrand $TargetBrand
        }
        
        Write-Host " Done" -ForegroundColor Green
        
        if ($SuppressPrompt) {
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $uniqueContacts.Count -Silent
        }
        else {
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $uniqueContacts.Count -NoPrompt
        }
        
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
    Write-Host "  Printer Address Book Converter v2.0" -ForegroundColor Cyan
    Write-Host "  Log file: $Script:LogFile" -ForegroundColor Gray
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-FunctionEntry -FunctionName 'Main' -Parameters @{ SourcePath = $SourcePath; TargetPath = $TargetPath; Mode = $Mode; NoInteractive = $NoInteractive }
    Write-Log -Level 'INFO' -Function 'Main' -Message "Session started"

    # Non-interactive mode
    if ($NoInteractive) {
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host "  NON-INTERACTIVE MODE" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""

        # Handle array of source paths (for Merge/Batch mode)
        if ($SourcePath -is [array]) {
            $sourceFiles = $SourcePath
        }
        else {
            $sourceFiles = @($SourcePath)
        }
        
        # Validate all source files exist
        $missingFiles = @()
        foreach ($file in $sourceFiles) {
            if (-not (Test-Path $file)) {
                $missingFiles += $file
            }
        }
        
        if ($missingFiles.Count -gt 0) {
            Write-Host "Error: Source file(s) not found:" -ForegroundColor Red
            foreach ($missing in $missingFiles) {
                Write-Host "  - $missing" -ForegroundColor Red
            }
            return
        }

        # Validate Target/Template
        $TemplateToUse = $null
        if ($TemplatePath -and (Test-Path $TemplatePath)) {
            $TemplateToUse = $TemplatePath
            Write-Host "Using template: $TemplateToUse" -ForegroundColor Green
        }
        elseif ($TargetPath -and (Test-Path $TargetPath)) {
            $TemplateToUse = $TargetPath
            Write-Host "Using target file as template: $TemplateToUse" -ForegroundColor Green
        }
        else {
            Write-Host "Error: Existing Target CSV file (Template) is required." -ForegroundColor Red
            return
        }
        
        # Analyze Template
        Write-Host "Analyzing template... " -NoNewline
        $templateDetails = Get-CsvDetails -FilePath $TemplateToUse
        $templateStructure = Get-CsvStructure -FilePath $TemplateToUse -Encoding $templateDetails.Encoding
        
        $colHeaderLine = $templateStructure.Headers | 
        Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_) -and 
            -not ($_ -match '^#') -and 
            -not ($_ -match '^"alternative"') -and
            -not ($_ -match '^@Ver')
        } | Select-Object -Last 1

        if ($templateStructure.HasQuotedHeaders) {
            $templateColumns = ($colHeaderLine -split $templateDetails.Delimiter).TrimStart('"').TrimEnd('"')
        }
        else {
            $templateColumns = $colHeaderLine -split $templateDetails.Delimiter
        }
        $templateColumns = $templateColumns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $templateMapping = Get-TemplateMapping -Headers $templateColumns
        Write-Host "OK" -ForegroundColor Green

        # -----------------------------------------------------
        # PATH 1: OUTLOOK MODE
        # -----------------------------------------------------
        if ($Mode -eq 'Outlook') {
            Write-Host "Mode: Outlook Import" -ForegroundColor Cyan
            $text = Get-Content $SourcePath -Raw
            $rawContacts = Parse-OutlookContacts -InputText $text
             
            if ($rawContacts.Count -eq 0) {
                Write-Host "Error: No valid contacts found in input text." -ForegroundColor Red
                exit 1
            }
             
            Write-Host "Parsing... $($rawContacts.Count) contacts found"
             
            $finalContacts = @()
            foreach ($rc in $rawContacts) {
                $norm = ConvertTo-NormalizedContact -Contact $rc -SourceMapping @{ Email = 'Email'; DisplayName = 'Name' }
                if ($norm) { $finalContacts += $norm }
            }
             
            $uniqueContacts = Remove-Duplicates -Contacts $finalContacts
            $outputPath = Get-SafeOutputPath -SourcePath $SourcePath -TargetBrand "Converted"
             
            Write-Host "Writing $($uniqueContacts.Count) unique contacts..." -NoNewline
            Write-AddressBook -NormalizedContacts $uniqueContacts -OutputPath $outputPath -TargetBrand "Converted" `
                -TemplateStructure $templateStructure -TemplateDetails $templateDetails -TemplateMapping $templateMapping -TemplateColumns $templateColumns
             
            Write-Host " Done" -ForegroundColor Green
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $uniqueContacts.Count -NoPrompt
            exit 0
        }
        
        # -----------------------------------------------------
        # PATH 2: MERGE MODE (Multiple files)
        # -----------------------------------------------------
        if ($Mode -eq 'Merge' -and $sourceFiles.Count -gt 1) {
            Write-Host "Mode: Merge $($sourceFiles.Count) files" -ForegroundColor Cyan
            
            $allContacts = @()
            foreach ($sourceFile in $sourceFiles) {
                # Detect brand
                $detection = Get-SourceBrand -FilePath $sourceFile
                if ($detection.Confidence -ne 100) {
                    Write-Host "Warning: Skipping $sourceFile (unknown format)" -ForegroundColor Yellow
                    continue
                }
                
                Write-Host "Loading $(Split-Path -Leaf $sourceFile) ($($detection.Brand))..." -NoNewline
                $contacts = Read-AddressBook -FilePath $sourceFile -Brand $detection.Brand
                $validContacts = Validate-Contacts -Contacts $contacts
                $allContacts += $validContacts
                Write-Host " $($validContacts.Count) contacts" -ForegroundColor Green
            }
            
            Write-Host ""
            Write-Host "Total loaded: $($allContacts.Count)" -ForegroundColor Cyan
            
            $uniqueContacts = Remove-Duplicates -Contacts $allContacts
            Write-Host "After deduplication: $($uniqueContacts.Count)" -ForegroundColor Green
            
            # Normalize for output
            $finalContacts = @()
            foreach ($c in $uniqueContacts) {
                $norm = ConvertTo-NormalizedContact -Contact $c -SourceMapping @{ Email = 'Email'; DisplayName = 'Name' }
                if ($norm) { $finalContacts += $norm }
            }
            
            $outputPath = Get-SafeOutputPath -SourcePath $sourceFiles[0] -TargetBrand "Merged" -IsMerge $true
            
            Write-Host "Writing merged output..." -NoNewline
            Write-AddressBook -NormalizedContacts $finalContacts -OutputPath $outputPath -TargetBrand "Converted" `
                -TemplateStructure $templateStructure -TemplateDetails $templateDetails -TemplateMapping $templateMapping -TemplateColumns $templateColumns
            
            Write-Host " Done" -ForegroundColor Green
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $finalContacts.Count -NoPrompt
            exit 0
        }
        
        # -----------------------------------------------------
        # PATH 3: CSV CONVERSION (Single or Batch)
        # -----------------------------------------------------
        if ($sourceFiles.Count -eq 1) {
            $SourcePath = $sourceFiles[0]
        }
        
        if (-not (Test-SafePath -Path $SourcePath)) {
            Write-Host "Error: Invalid file path" -ForegroundColor Red
            return
        }

        # Detect source brand
        $detection = Get-SourceBrand -FilePath $SourcePath
        if ($detection.Confidence -ne 100) {
            Write-Host "Error: Source format detection failed." -ForegroundColor Red
            return
        }

        $sourceBrand = $detection.Brand
        Write-Host "Detected source: $sourceBrand" -ForegroundColor Green
        
        Reset-Stats

        $converted = Convert-AddressBook -SourcePath $SourcePath -SourceBrand $sourceBrand -TargetBrand "Converted" -TargetTemplatePath $TemplateToUse

        if ($converted) {
            Show-ValidationReport
            Write-Log "INFO" "Session completed (non-interactive)"
            Write-Host ""
            exit 0
        }
        else {
            Write-Host "Conversion failed" -ForegroundColor Red
            Write-Log "ERROR" "Session failed (non-interactive)"
            Write-Host ""
            exit 1
        }
    }

    # Interactive mode (original behavior)
    $modeOptions = @(
        'Convert Files (Single or Multiple)',
        'Merge Multiple Files into One',
        'Create from Outlook list',
        'Exit'
    )
    $desc = "Auto-detects source and target formats from existing CSV files.`nConverted files saved to same directory as source.`n`n⚠️  IMPORTANT: All CSV files must contain at least 1 contact (required for format detection)"
    $modeIndex = Show-Menu -Options $modeOptions -Title "Select operation:" -Description $desc

    if ($modeIndex -eq -1 -or $modeIndex -eq 3) {
        Write-Host ""
        Write-Host "Cancelled" -ForegroundColor Yellow
        Write-Host ""
        return
    }

    $mode = switch ($modeIndex) {
        0 { 'Convert' }
        1 { 'Merge' }
        2 { 'Outlook' }
    }

    $sourceFiles = @()

    if ($mode -eq 'Convert') {
        # Allow selecting one or multiple files
        $sourceFile = Select-Files -MultiSelect $true
        if (-not $sourceFile) {
            Write-Host ""
            Write-Host "No file selected" -ForegroundColor Yellow
            Write-Host ""
            return
        }
        $sourceFiles += $sourceFile
    }
    elseif ($mode -eq 'Merge') {
        $selectedFiles = Select-Files -MultiSelect $true
        if (-not $selectedFiles -or $selectedFiles.Count -eq 0) {
            Write-Host ""
            Write-Host "No files selected" -ForegroundColor Yellow
            Write-Host ""
            return
        }
        $sourceFiles = $selectedFiles
    }
    elseif ($mode -eq 'Outlook') {
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host "  CREATE ADDRESS BOOK FROM OUTLOOK LIST" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Instructions:" -ForegroundColor Yellow
        Write-Host "1. In Outlook, paste email addresses in To: field"
        Write-Host "2. Press Ctrl+K (Check Names) to format properly"
        Write-Host "3. Copy the formatted result and paste here"
        Write-Host ""
        Write-Host "Expected format: LASTNAME, Firstname <email@domain.org>; ..."
        Write-Host ""
        Write-Host "Paste contact list below (press Enter twice when done):" -ForegroundColor Green
        Write-Host ""

        $inputLines = @()
        while ($true) {
            $line = Read-Host
            if ([string]::IsNullOrWhiteSpace($line)) { break }
            $inputLines += $line
        }

        if ($inputLines.Count -eq 0) {
            Write-Host ""
            Write-Host "No input provided" -ForegroundColor Yellow
            Write-Host ""
            return
        }

        # Parse Outlook format
        Write-Host ""
        Write-Host "Parsing contacts..." -ForegroundColor Cyan
        
        $contactsRaw = Parse-OutlookContacts -InputText ($inputLines -join "`r`n")
        
        if ($contactsRaw.Count -eq 0) {
            # ...
        }

        # Normalize interactive
        $contacts = @()
        foreach ($c in $contactsRaw) {
            $norm = ConvertTo-NormalizedContact -Contact $c -SourceMapping @{ Email = 'Email'; DisplayName = 'Name' }
            if ($norm) { $contacts += $norm }
        }

        if ($contacts.Count -eq 0) {
            Write-Host ""
            Write-Host "No valid contacts found. Check format and try again." -ForegroundColor Red
            Write-Host ""
            return
        }

        Write-Host ""
        Write-Host "Parsed $($contacts.Count) contacts:" -ForegroundColor Green
        # $contacts | ForEach-Object { Write-Host "  $($_.DisplayName) - $($_.Email)" } # Too many logs

        # Select target TEMPLATE
        Write-Host ""
        Write-Host "Please select a Target CSV Template file..." -ForegroundColor Cyan
        $targetFile = Select-Files -MultiSelect $false

        if (-not $targetFile) {
            Write-Host ""
            Write-Host "Cancelled" -ForegroundColor Yellow
            Write-Host ""
            return
        }
        
        # Analyze Template
        Write-Host "Analyzing template... " -NoNewline
        $templateDetails = Get-CsvDetails -FilePath $targetFile
        $templateStructure = Get-CsvStructure -FilePath $targetFile -Encoding $templateDetails.Encoding
        
        $colHeaderLine = $templateStructure.Headers | 
        Where-Object { 
            -not [string]::IsNullOrWhiteSpace($_) -and 
            -not ($_ -match '^#') -and 
            -not ($_ -match '^"alternative"') -and
            -not ($_ -match '^@Ver')
        } | Select-Object -Last 1

        if ($templateStructure.HasQuotedHeaders) {
            $templateColumns = ($colHeaderLine -split $templateDetails.Delimiter).TrimStart('"').TrimEnd('"')
        }
        else {
            $templateColumns = $colHeaderLine -split $templateDetails.Delimiter
        }
        $templateColumns = $templateColumns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $templateMapping = Get-TemplateMapping -Headers $templateColumns
        Write-Host "OK" -ForegroundColor Green

        # Generate output file - save to current directory
        $outputDir = Get-Location
        $timestamp = Get-Date -Format 'yyyy-MM-dd'
        $outputPath = Join-Path $outputDir "Outlook_converted_${timestamp}.csv"

        # Validate and deduplicate
        Reset-Stats
        $validContacts = Validate-Contacts -Contacts $contacts
        $uniqueContacts = Remove-Duplicates -Contacts $validContacts

        Write-Host ""
        Write-Host "Writing output..." -NoNewline
        try {
            Write-AddressBook -NormalizedContacts $uniqueContacts -OutputPath $outputPath -TargetBrand "Converted" `
                -TemplateStructure $templateStructure -TemplateDetails $templateDetails -TemplateMapping $templateMapping -TemplateColumns $templateColumns
                
            Write-Host " Done" -ForegroundColor Green
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $uniqueContacts.Count
        }
        catch {
            Write-Host " FAILED" -ForegroundColor Red
            Write-Host ""
            Write-Host "Error: $_" -ForegroundColor Red
            Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        }

        Show-ValidationReport
        Write-Log "INFO" "Session completed (Outlook mode)"
        Write-Host ""
        return
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
    Write-Host "Select target CSV file (format will be auto-detected)..." -ForegroundColor Cyan
    $targetFile = Select-Files -MultiSelect $false
    if (-not $targetFile) {
        Write-Host ""
        Write-Host "No target file selected" -ForegroundColor Yellow
        Write-Host ""
        return
    }

    if (-not (Test-SafePath -Path $targetFile)) {
        Write-Host ""
        Write-Host "Error: Invalid target file path" -ForegroundColor Red
        Write-Host ""
        return
    }

    # Use as Template
    Write-Host "Using as Template: " -NoNewline
    Write-Host (Split-Path -Leaf $targetFile) -ForegroundColor Green
    $targetTemplatePath = $targetFile
    $targetBrand = "Converted" # Label for filename generation

    Write-Log "INFO" "Target Template: $targetFile"

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
            # Pass TemplatePath logic if needed for merge, but usually Merge uses Convert-AddressBook logic internally? 
            # Wait, Write-AddressBook is called directly here. 
            # We need to analyze template here for Merge mode too, or refactor Merge to use Convert logic.
            # Simplify: We'll use the same robust Template analysis logic that Convert-AddressBook uses.
            
            # 1. Analyze Template
            $templateDetails = Get-CsvDetails -FilePath $targetTemplatePath
            $templateStructure = Get-CsvStructure -FilePath $targetTemplatePath -Encoding $templateDetails.Encoding
            
            # Find header
            $colHeaderLine = $templateStructure.Headers | 
            Where-Object { 
                -not [string]::IsNullOrWhiteSpace($_) -and 
                -not ($_ -match '^#') -and 
                -not ($_ -match '^"alternative"') -and
                -not ($_ -match '^@Ver')
            } | Select-Object -Last 1
                
            if ($templateStructure.HasQuotedHeaders) {
                $templateColumns = ($colHeaderLine -split $templateDetails.Delimiter).TrimStart('"').TrimEnd('"')
            }
            else {
                $templateColumns = $colHeaderLine -split $templateDetails.Delimiter
            }
            $templateColumns = $templateColumns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $templateMapping = Get-TemplateMapping -Headers $templateColumns

            $outputPath = Get-SafeOutputPath -SourcePath $fileInfo[0].Path -TargetBrand "Merged" -IsMerge $true

            Write-Host ""
            Write-Host "  Writing merged output..." -NoNewline
            
            # Map valid contacts to normalized structure (they are already normalized by Validate-Contacts/Read-AddressBook?)
            # Read-AddressBook returns object with Name/Email. ConvertTo-NormalizedContact does Normalize.
            # Convert-AddressBook does Normalized.
            # Merge loop above calls Read-AddressBook which returns basic objects. 
            # We need to Normalize them properly before writing if we want hygiene.
            
            # FIX: Re-normalize merged contacts to ensure capitalization etc
            $finalContacts = @()
            foreach ($c in $uniqueContacts) {
                # Convert basic object to Normalized
                $norm = ConvertTo-NormalizedContact -Contact $c -SourceMapping @{ Email = 'Email'; DisplayName = 'Name' }
                if ($norm) { $finalContacts += $norm }
            }

            Write-AddressBook -NormalizedContacts $finalContacts -OutputPath $outputPath -TargetBrand "Converted" `
                -TemplateStructure $templateStructure -TemplateDetails $templateDetails -TemplateMapping $templateMapping -TemplateColumns $templateColumns

            Write-Host " Done" -ForegroundColor Green
            Show-OutputSuccess -OutputPath $outputPath -ContactCount $finalContacts.Count
        }
        catch {
            Write-Host " FAILED" -ForegroundColor Red
            Write-Host "  Error: $_" -ForegroundColor Red
        }
    }
    else {
        $fileCount = 1
        $totalFiles = $fileInfo.Count
        $outputPaths = @()
        $isMultiFile = $totalFiles -gt 1

        foreach ($info in $fileInfo) {
            if ($isMultiFile) {
                Write-Host ""
                Write-Host "[File $fileCount of $totalFiles]" -ForegroundColor Gray
            }

            $converted = Convert-AddressBook -SourcePath $info.Path -SourceBrand $info.Brand -TargetTemplatePath $targetTemplatePath -SuppressPrompt:$isMultiFile

            if ($converted) {
                $allConvertedContacts += $converted
                
                # Track output path for multi-file batch
                if ($isMultiFile) {
                    $outputPath = Get-SafeOutputPath -SourcePath $info.Path -TargetBrand "Converted"
                    $outputPaths += $outputPath
                }
            }

            $fileCount++
        }
        
        # Show unified success message for multi-file batch
        if ($isMultiFile -and $outputPaths.Count -gt 0) {
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "SUCCESS! $($outputPaths.Count) files converted:" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "Location: " -NoNewline -ForegroundColor White
            $outputDir = Split-Path $outputPaths[0] -Parent
            Write-Host $outputDir -ForegroundColor Cyan
            Write-Host ""
            Write-Host "Files:" -ForegroundColor Gray
            foreach ($path in $outputPaths) {
                $fileName = Split-Path $path -Leaf
                Write-Host "  - $fileName" -ForegroundColor Gray
            }
            Write-Host ""
            
            # Offer to open folder once
            $openFolder = Read-Host "Open folder in Explorer? (Y/N)"
            if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
                Start-Process "explorer.exe" -ArgumentList "`"$outputDir`""
            }
        }
    }

    Show-ValidationReport

    Write-Log -Level 'INFO' -Function 'Main' -Message "Session completed"
    Write-FunctionExit -FunctionName 'Main'
    Write-Host ""
}

Main

#endregion
