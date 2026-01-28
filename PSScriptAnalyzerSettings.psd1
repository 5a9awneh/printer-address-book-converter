@{
    # Exclude rules for interactive scripts
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'  # Intentional: Interactive UI requires Write-Host for colored output
    )
    
    # Include specific rules to check
    IncludeRules = @(
        'PSAvoidUsingEmptyCatchBlock'
    )
    
    # Rules to run
    Rules = @{
        PSAvoidUsingEmptyCatchBlock = @{
            Enable = $true
        }
        PSUseSingularNouns = @{
            Enable = $false  # Disable for functions like Remove-Duplicates, Validate-Contacts
        }
    }
}
