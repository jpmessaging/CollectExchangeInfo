function Load-CliXml {
    #requires -Version 3.0
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        [Switch]$SkipVariableCreation
    )

    $resolvedPath = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (!$resolvedPath) {
        throw "$Path doesn't exit"
    }

    Write-Verbose "Path is $resolvedPath"

    # Process each file
    $sw = [System.Diagnostics.StopWatch]::StartNew()
    $files = Get-ChildItem $resolvedPath -Filter "*.xml"

    $importedObjects = New-Object System.Collections.Generic.List[object]
    $importedFileCount = 0
    $failedFiles = @(
        foreach ($file in $files) {
            try {
                $importedValue = Import-Clixml -Path $file.FullName -ErrorAction:Stop
                $importedFileCount++

                # if SkipVariableCreation, then imported objects are outputted as 'ImportedObjects'.
                if ($SkipVariableCreation) {
                    $importedObjects.Add($importedValue)
                }
                else {
                    # Check if the variable exist already
                    $variableName = $file.BaseName
                    $variable = Get-Variable $variableName -Scope Global -ErrorAction SilentlyContinue

                    # if already exist, remove it.
                    if ($variable) {
                        Write-Verbose "Removing $variableName"
                        Remove-Variable $variableName -Scope Global
                    }

                    # Create a variable in Global Scope and set the value
                    New-Variable -Name $variableName -Scope Global -Value $importedValue
                    Write-Verbose "$variableName is created"
                    Write-Progress -Activity "Loading CliXml" -Status "$importedFileCount/$($files.Count) files loaded." -PercentComplete:($importedFileCount/$files.Count*100)
                }
            }
            catch [System.Xml.XmlException],[System.Security.Cryptography.CryptographicException] {
                Write-Verbose "Cannot import $($file.FullName)"
                Write-Output $($file.FullName)
            }
        }
        Write-Progress -Activity "Loading CliXml" -Completed
    )

    $sw.Stop()

    if ($SkipVariableCreation -and ($importedFileCount -ne $importedObjects.Count)) {
        Write-Warning "[BUG] importedFileCount ($importedFileCount) doesn't match importedObjects.Count ($($importedObjects.Count))."
    }

    [PSCustomObject]@{
        VariablesCreated = !$SkipVariableCreation;
        ImportedFileCount = $importedFileCount;
        ImportedObjects = $importedObjects;
        FailedFileCount = $failedFiles.Count;
        FailedFiles = $failedFiles;
        Elapsed = $sw.Elapsed
    }
}
