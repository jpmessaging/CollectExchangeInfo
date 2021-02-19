<#
.DESCRIPTION
Load all CliXml files in a folder (or subfolders when Recurse is used).
#>
function Load-CliXml {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        [switch]$Recurse,
        [switch]$SkipVariableCreation,
        # Remove all the variables previously created
        [switch]$ClearPreviousDataSet
    )

    $resolvedPath = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (!$resolvedPath) {
        Write-Error "$Path doesn't exit"
        return
    }

    # If variable 'LoadCliXmlData' already exits, it means Load-CliXmls was run before.
    # Remove all the variables previously created.
    if ($ClearPreviousDataset) {
        foreach ($v in $LoadCliXmlData.Variables) {
            Remove-Variable $v -Scope Global -ErrorAction SilentlyContinue
        }
        if ($LoadCliXmlData.Variables) {
            $LoadCliXmlData.Variables.Clear()
        }
    }

    $sw = [System.Diagnostics.StopWatch]::StartNew()
    $files = Get-ChildItem $resolvedPath -Filter:"*.xml" -Recurse:$Recurse.IsPresent
    $importedFileCount = 0

    if ($SkipVariableCreation) {
        # Objects will be stored ad {name, object}
        $importedObjects = New-Object System.Collections.Generic.Dictionary[[string]`,[PSObject]]
    }

    if ($null -eq $LoadCliXmlData.Variables) {
        $varNames = New-Object System.Collections.Generic.List[string]
    }
    else {
        $varNames = $LoadCliXmlData.Variables
    }

    $failedFiles = @(
        foreach ($file in $files) {
            try {
                # Do some basic check if the xml is a CliXml (not random xml file)
                # The first line of CliXml should look like "<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">""
                if ((Get-Content -Path $file.FullName -TotalCount 1) -notlike '<Objs*' ) {
                    # This does not look like a CliXml. Skip it.
                    continue
                }

                $importedValue = Import-Clixml -Path $file.FullName -ErrorAction:Stop
                $importedFileCount++

                # if SkipVariableCreation, then imported objects are outputted as 'ImportedObjects'.
                if ($SkipVariableCreation) {
                    if (-not $importedObjects.ContainsKey($file.BaseName)) {
                        $importedObjects.Add($file.BaseName, $importedValue)
                    }
                    else {
                        Write-Error "There are files with the same name `"$($file.BaseName)`"."
                    }
                }
                else {
                    $variableName = $file.BaseName
                    # Write a warning if overwriting an existing variable.
                    if ($LoadCliXmlData.Variables -and $LoadCliXmlData.Variables.Contains($variableName)) {
                        Write-Warning "Overwriting variable `"$variableName`". Use -ClearPreviousDataSet to clear previous data set."
                    }

                    # Create a variable in Global Scope and set the value. -Force will overwrite the existing one if any.
                    New-Variable -Name $variableName -Scope Global -Value $importedValue -Force

                    # Remember the name of variables created so that they can be removed later when this function is called later with ClearPreviousDataSet.
                    $varNames.Add($variableName)
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

    # Add "LoadCliXmlData" as a Script-scoped variable.
    # LoadCliXmlData keeps "Path" of the most recent run and names of all the variables it created (including duplicates).
    Remove-Variable 'LoadCliXmlData' -Scope Script -ErrorAction SilentlyContinue
    New-Variable -Name 'LoadCliXmlData' -Scope Script -Value $([PSCustomObject]@{Path = $Path; Variables = $varNames})

    if ($SkipVariableCreation -and ($importedFileCount -ne $importedObjects.Count)) {
        Write-Warning "[BUG] importedFileCount ($importedFileCount) doesn't match importedObjects.Count ($($importedObjects.Count))."
    }

    [PSCustomObject]@{
        ImportedFileCount = $importedFileCount
        ImportedObjects = $importedObjects
        FailedFileCount = $failedFiles.Count;
        FailedFiles = $failedFiles
        Elapsed = $sw.Elapsed
    }
}