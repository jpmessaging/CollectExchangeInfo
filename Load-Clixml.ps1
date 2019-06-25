<#
  Load all Clixml files in a folder
  if Path is not specified, it uses the current directory
#>
function Load-Clixml
{
    #requires -Version 3.0
    [CmdletBinding()]
    param
    (
    [Parameter(Mandatory=$true)]
    $Path,
    [Switch]$SkipVariableCreation
    )

    # Resolve Path
    $resolvedPath = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (!$resolvedPath)
    {        
        Write-Error "$Path doesn't exit"
        return
    }

    Write-Verbose "Path is $resolvedPath"

    # Process each file
    $sw = [System.Diagnostics.StopWatch]::StartNew()
    $files = Get-ChildItem $resolvedPath -Filter "*.xml"    
    
    $importedObjects = New-Object System.Collections.Generic.List[object]
    $importedFileCount = 0
    $failedFiles = @(
        foreach ($file in $files)
        {                    
            try
            {
                $importedValue = Import-Clixml -Path $file.FullName -ErrorAction:Stop
                $importedFileCount++

                # if SkipVariableCreation, then imported objects are outputted as 'ImportedObjects'.
                if ($SkipVariableCreation)
                {                    
                    $importedObjects.Add($importedValue)                    
                }
                else
                {
                    # Check if the variable exist already
                    $variableName = $file.BaseName
                    $variable = Get-Variable $variableName -Scope Global -ErrorAction SilentlyContinue

                    # if already exist, remove it.
                    if ($variable)
                    {
                        Write-Verbose "Removing $variableName"
                        Remove-Variable $variableName -Scope Global                
                    }

                    # Create a variable in Global Scope and set the value
                    New-Variable -Name $variableName -Scope Global -Value $importedValue
                    Write-Verbose "$variableName is created"
                }
            }
            catch [System.Xml.XmlException],[System.Security.Cryptography.CryptographicException] 
            {
                Write-Verbose "Cannot import $($file.FullName)"
                Write-Output $($file.FullName)
            }
        }
    )

    $sw.Stop()
    
    if ($SkipVariableCreation -and ($importedFileCount -ne $importedObjects.Count))
    {
        Write-Warning "[BUG] importedFileCount ($importedFileCount) doesn't match importedObjects.Count ($($importedObjects.Count))."
    }

    Write-Output (
        [PSCustomObject]@{VariablesCreated = !$SkipVariableCreation; ImportedFileCount = $importedFileCount; ImportedObjects = $importedObjects; FailedFileCount = $failedFiles.Count ;FailedFiles = $failedFiles; Elapsed = $sw.Elapsed}
    )
}
