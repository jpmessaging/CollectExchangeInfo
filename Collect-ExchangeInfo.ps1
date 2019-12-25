<#
.SYNOPSIS
Collect Exchange Related Information

.DESCRIPTION
This script runs various Exchange Server cmdlets (and some other cmdlets) to collect information.
It runs only Get-* or Test-* cmdlets and doesn't make any changes.
Collected information will be saved in the folder specified with Path parameter.

This script MUST be ran with PowerShell v2 or above (#requires checks this)
Running with Exchange Management Shell (EMS) & Manual EMS are supported.

   What I mean by Manual EMS is:
   $Server = "myExchange.contoso.com"
   $connectionUrl = "http://" + $Server + "/powershell/?serializationLevel=full"
   Import-PSSession (New-PSSession -ConnectionUri $connectionUrl -ConfigurationName Microsoft.Exchange)

.PARAMETER Path
Mandatory path for output files.
This can be either absolute or relative path.
If the specified path doesn't exist, it will be created.

.PARAMETER Parameters
List of global parameters.
These parameters are applied whenever available for a cmdlet.
For example, if DomainContoller:dc.contoso.local is specified, this parameter will be specified for every cmdlet which supports -DomainContoller.
Notice that you don't need a hyphen ('-') in front of parameter name.

.PARAMETER Servers
List of servers to directly access to.
The servers listed here will be directly touched for some of the cmdlets such as Get-*VirtualDirectory.
Wild card is supported. For example, "E201*" includes all the servers whose name matches "E201*"

If not specified, no servers are directly accessed and only Active Directory information is gathered (For exampple, Get-*VirtualDirectory will be run with -ADPropertiesOnly).

Note: Connectivity for each server is checked by ping before running any Exchange cmdlets.
If a server fails on connectivity test, it won't be accessed for the rest of execution.

.PARAMETER IncludeFIPS
Switch to include FIPS (Forefront Information Protection Service) related information.

.PARAMETER IncludeEventLogs
Switch to include Application & System event logs on the servers specified in "Servers" parameter.

.PARAMETER IncludeEventLogsWithCrimson
Switch to include Exchange-related Crimson logs ("Microsoft-Exchange-*") as well as Application & System event logs on the servers specified in "Servers" parameter.

.PARAMETER IncludePerformanceLog
Switch to include Exchange's Perfmon log (Only Exchange 2013 and above collects perfmon log by default)

.PARAMETER KeepOutputFiles
Switch to keep the output files. If this is not specified, all the output files will be deleted after being packed to a zip file.
In order to avoid deleting unrelated files or folders, this script makes sure that the folder specified by Path paramter is empty and if not empty, it stops executing.

.EXAMPLE
.\Collect-ExchangeInfo -Path .\exinfo -Servers:*

Create (if not exist) a sub folder "exinfo" under the current path.
All the output files are saved in this folder.
All Exchange Servers will be accessed since * is specified for "Servers".

Note that running on Exchange 2010 will NOT find Exchange 2013 & 2016 servers.  So It's recommended to run on the latest version of Exchange Server in the organization.

.EXAMPLE
.\Collect-ExchangeInfo -Path C:\exinfo

Create (if not exist) C:\exinfo and save output files there.
No servers are accessed since Servers parameter is not specified (i.e. Only information from Active Directly is collected)

.EXAMPLE
.\Collect-ExchangeInfo -Path C:\exinfo -Servers:EX-* -IncludeEventLogsWithCrimson

Create (if not exist) C:\exinfo and save output files there.
Exchange Servers matching "EX-*" will be directly accessed and their event logs including Exchange's crimson logs will be collected.

.NOTES
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [string[]]$Parameters = @(),
    [string[]]$Servers = @(),
    [switch]$IncludeFIPS,
    [switch]$IncludeEventLogs = $false,
    [switch]$IncludeEventLogsWithCrimson,
    [switch]$IncludeIISVirtualDirectories,
    [switch]$IncludePerformanceLog,
    [switch]$KeepOutputFiles
)

$version = "2019-12-25"
#requires -Version 2.0

<#
  Save object(s) to a text file and optionally export to CliXml.
#>
function Save-Object {
    [CmdletBinding()]
    Param(
        #[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Parameter(ValueFromPipeline=$true)]
        $object,
        $Name,
        [string]$Path = $Script:Path,
        [bool]$WithCliXml = $true,
        $Depth = 5 # depth for Export-CliXml
    )

    begin {
        # Need to accumulate result to support pipeline. Use List<> to improve performance
        $objectList = New-Object System.Collections.Generic.List[object]
        [string]$objectName = $Name
    }

    process {
        # Validate the given objects.  If valid, collect them in a list.
        # Collected objects are outputted in the END block

        # When explicitly passed, object is actually a list of objects.
        # When passed from pipeline, object is a single object.
        # To deal with this, use foreach.

        foreach ($o in $object) {
            if ($null -eq $o) {
                return
            }
            <#
            elseif($o -is [string])
            {
                # assume a string object is an error and write it to log
                Write-Log $o
            }
            #>
            else {
                if (-not($objectName)) {
                    $objectName = $o.GetType().Name
                }
                $objectList.Add($o)
            }
        }
    }

    end {
        if ($objectList.Count -gt 0) {
            if(-not $objectName) {
                Write-Log "[$($MyInvocation.MyCommand)] Error:objectName is null"
            }

            if ($WithCliXml) {
                $objectList | Export-Clixml -Path:([System.IO.Path]::Combine($Path, "$objectName.xml")) -Encoding:UTF8 -Depth $Depth
            }

            $objectList | Format-List * | Out-File ([System.IO.Path]::Combine($Path, "$objectName.txt")) -Encoding:UTF8
        }
    }
}

<#
  Compress a folder and create a zip file.
#>
function Compress-Folder {
    [CmdletBinding()]
    param(
        # Specifies a path to one or more locations.
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$Destination,
        [string]$ZipFileName,
        [switch]$IncludeDateTime,
        [switch]$RemoveFiles,
        [switch]$UseShellApplication
    )

    $Path = Resolve-Path $Path
    $zipFileNameWithouExt = [System.IO.Path]::GetFileNameWithoutExtension($ZipFileName)
    if ($IncludeDateTime) {
        $zipFileName = $zipFileNameWithouExt + "_" + "$(Get-Date -Format "yyyyMMdd_HHmmss").zip"
    }
    else {
        $zipFileName = "$zipFileNameWithouExt.zip"
    }

    # If Destination is not given, use %TEMP% folder.
    if (-not $Destination) {
        $Destination = $env:TEMP
    }

    if (-not (Test-Path $Destination)) {
        New-Item $Destination -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Destination = Resolve-Path $Destination
    $zipFilePath = Join-Path $Destination -ChildPath $zipFileName

    $NETFileSystemAvailable = $false

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        # Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $NETFileSystemAvailable = $true
    }
    catch {
        Write-Warning "System.IO.Compression.FileSystem wasn't found. Using alternate method"
    }

    if ($NETFileSystemAvailable -and $UseShellApplication -eq $false) {
        # Note: [System.IO.Compression.ZipFile]::CreateFromDirectory() fails when one or more files in the directory is locked.
        #[System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $zipFilePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

        try {
            New-Item $zipFilePath -ItemType file | Out-Null

            $zipStream = New-Object System.IO.FileStream -ArgumentList $zipFilePath, ([IO.FileMode]::Open)
            $zipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipStream, ([IO.Compression.ZipArchiveMode]::Create)

            $files = @(Get-ChildItem $Path -Recurse | Where-Object {-not $_.PSIsContainer})
            $count = 0

            foreach ($file in $files) {
                Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Adding $($file.FullName)" -PercentComplete (100 * $count / $files.Count)

                try {
                    $fileStream = New-Object System.IO.FileStream -ArgumentList $file.FullName, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::ReadWrite)
                    $zipEntry = $zipArchive.CreateEntry($file.FullName.Substring($Path.Length + 1))
                    $zipEntryStream = $zipEntry.Open()
                    $fileStream.CopyTo($zipEntryStream)

                    ++$count
                }
                catch {
                    Write-Error "Failed to add $($file.FullName). $_"
                }
                finally {
                    if ($fileStream) {
                        $fileStream.Dispose()
                    }

                    if ($zipEntryStream) {
                        $zipEntryStream.Dispose()
                    }
                }
            }
        }
        finally {
            if ($zipArchive) {
                $zipArchive.Dispose()
            }

            if ($zipStream) {
                $zipStream.Dispose()
            }

            Write-Progress -Activity "Creating a zip file $zipFilePath" -Completed
        }
    }
    else {
        # Use Shell.Application COM

        # Create a zip file manually
        $shellApp = New-Object -ComObject Shell.Application
        Set-Content $zipFilePath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (Get-Item $zipFilePath).IsReadOnly = $false

        $zipFile = $shellApp.NameSpace($zipFilePath)

        # If target folder is empty, CopyHere() fails. So make sure it's not empty
        if (@(Get-ChildItem $Path).Count -gt 0) {
            # Start copying the whole and wait until it's done. CopyHere works asynchronously.
            $zipFile.CopyHere($Path)

            # Now wait and poll
            $inProgress = $true
            $delayMilliseconds = 200
            Start-Sleep -Milliseconds 3000
            [System.IO.FileStream]$file = $null
            while ($inProgress) {
                Start-Sleep -Milliseconds $delayMilliseconds

                try {
                    $file = [System.IO.File]::Open($zipFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                    $inProgress = $false
                }
                catch [System.IO.IOException] {
                    Write-Debug $_.Exception.Message
                }
                finally {
                    if ($file) {
                        $file.Close()
                    }
                }
            }
        }
    }

    if (Test-Path $zipFilePath) {
        # If requested, remove zipped files
        if ($RemoveFiles) {
            Write-Verbose "Removing zipped files"
            Get-ChildItem $Path -Exclude $ZipFileName | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            $filesRemoved = $true
        }

        New-Object PSCustomObject -Property @{
            ZipFilePath = $zipFilePath.ToString()
            FilesRemoved = $filesRemoved -eq $true
        }
    }
    else {
        throw "Zip file wasn't successfully created at $zipFilePath"
    }
}

<#
  Runs Ldifde for Exchange organization in configuration context
#>
function Invoke-Ldifde {
    param (
      [Parameter(Mandatory=$true)]
      [string]$Path,
      [string]$FileName = "Ldifde.txt"
    )

    # if Path doesn't exit, create it
    if (-not (Test-Path $Path)) {
       New-Item -ItemType directory $Path | Out-Null
    }

    $resolvedPath  = Resolve-Path $Path -ErrorAction SilentlyContinue
    $filePath = Join-Path -Path $resolvedPath -ChildPath $FileName

    # Check if Ldifde.exe exists
    $isLdifdeAvailable = $false
    foreach ($path in $env:Path.Split(";")) {
        if ($path) {
            $exePath = Join-Path -Path $Path -ChildPath "ldifde.exe"
            if (Test-Path $exePath) {
                $IsLdifdeAvailable = $true;
                break;
            }
        }
    }

    if (-not $isLdifdeAvailable ) {
        throw "Ldifde is not available"
    }

    $exorg = (Get-OrganizationConfig).DistinguishedName

    if (-not $exorg) {
        throw "Couldn't get Exchange org DN"
    }

    # If this is an Edge server, use a port 50389.
    $server = Get-ExchangeServer $env:COMPUTERNAME
    if ($server -and $server.IsEdgeServer) {
        $Port = 50389
    }

    try {
        $fileNameWihtoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $stdOutput = Join-Path $resolvedPath -ChildPath "$fileNameWihtoutExtension.out"

        if ($Port) {
    		    $process = Start-Process ldifde -ArgumentList "-u -d `"$exorg`" -s localhost -t $Port -f `"$filePath`"" -PassThru -NoNewWindow -RedirectStandardOutput:$stdOutput
        }
        else {
            $process = Start-Process ldifde -ArgumentList "-u -d `"$exorg`" -f `"$filePath`"" -PassThru -NoNewWindow -RedirectStandardOutput:$stdOutput
        }

        if (-not $process.HasExited) {
            Wait-Process -InputObject $process
        }

        $process = $null
    }
    finally {
        if ($process) {
            Stop-Process -InputObject:$process -Force
            throw "ldifde was cancelled"
        }
    }
}

<#
  Run a given command only if it's available
  Run with parameters specified as Global Parameter (i.e. $script:Parameters)
#>
function RunCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command
    )

    $endOfCmdlet = $Command.IndexOf(" ")
    if ($endOfCmdlet -lt 0) {
        $cmdlet = $Command
    }
    else {
        $cmdlet = $Command.Substring(0, $endOfCmdlet)
    }

    # check if cmdlet is available
    $cmd = Get-Command $cmdlet -ErrorAction:SilentlyContinue
    if ($null -eq $cmd) {
        Write-Log "$cmdlet is not available"
        return
    }

    # check params
    # if any explicitly-requested params are not available, bail
    $paramMatches = Select-String " -(?<paramName>\w+)" -Input $Command -AllMatches

    if ($paramMatches) {
        $params = @(
            foreach($paramMatch in $paramMatches.Matches) {
                $paramName = $paramMatch.Groups['paramName'].Value

                # In order to support non-exact match, check each key
                $keyMatch = @(
                    foreach ($key in $cmd.Parameters.keys) {
                        if ($key -like "$($paramName)*") {
                            $key
                        }
                    }
                )

                # if there's no match or too many matches, bail.
                if ($keyMatch.Count -eq 0) {
                    Write-Log "Parameter '$paramName' is not available for $cmdlet"
                    return
                }
                elseif ($keyMatch.Count -gt 1) {
                    Write-Log "Parameter '$paramName' is ambiguous for $cmdlet"
                    return
                }

                $keyMatch[0]
            }
        )
    }

    # check if any parameter is requested globally
    # it's ok if these parameters are not available.
    foreach ($param in $script:Parameters) {
        $paramName = ($param -split ":")[0]

        if ($cmd.Parameters[$paramName]) {
            # explicitly-requested params take precedence
            # if not already in the list, add it.
            if ($params -notcontains $paramName) {
                $Command += " -$param"
           }
        }
    }

    # Finally run the command
    Write-Log "Running $Command"
    try {
        # capture non-terminating error
        $err = $($o = Invoke-Expression $Command) 2>&1
        if ($err) {
            Write-Log "[Non-Terminating Error] Error in '$Command'. $err $(if ($err.Exception.Line) {"(At line:$($err.Exception.Line) char:$($err.Exception.Offset))"})"
        }

        if ($null -ne $o) {
            Write-Output $o
        }
    }
    catch {
        # log terminating error.
        Write-Log "[Terminating Error] '$Command' failed. $_ $(if ($_.Exception.Line) {"(At line:$($_.Exception.Line) char:$($_.Exception.Offset))"})"
        if ($null -ne $Script:errs) {$errs.Add($_)}
    }
}

<#
  Run command against servers
#>
function Run {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        [string[]]$Servers,
        [string]$Identifier = "Server",
        [switch]$SkipIfNoServers,
        [Parameter(ValueFromPipeline=$true)]
        [object[]]$ResultCollection,
        [switch]$RemoveDuplicate,
        [switch]$PassThru
    )

    begin {
        $result = New-Object System.Collections.Generic.List[object]
    }
    # Accumulate the previous results
    process {
        # Make sure not to add $null and collection itself
        foreach ($pipedObj in $ResultCollection) {
            # In PowerShellV2, $null is iterated over.
            if ($pipedObj) {
                $result.Add($pipedObj)
            }
        }
    }

    end {
        $temp = @(
            if (-not $Servers.Count -and -not $SkipIfNoServers) {
                RunCommand $Command
            }
            elseif ($Servers) {
                foreach ($Server in $Servers) {
                    $firstTimeAddingServerName = $true
                    foreach ($entry in @(RunCommand "$Command -$Identifier $Server")) {
                        # Add ServerName prop if not exist already (but log only the first time per cmdlet)
                        if (!$entry.ServerName -and !$entry.Server -and !$entry.ServerFqdn -and !$entry.MailboxServer -and !$entry.Fqdn) {
                            if ($firstTimeAddingServerName) {
                                Write-Log "Adding ServerName to the result of '$Command -$Identifier $Server'"
                                $firstTimeAddingServerName = $false
                            }
                            # This is for PowerShell V2
                            # $entry | Add-Member -Type NoteProperty -Name:ServerName -Value:$Server
                            $entry = $entry | Select-Object *, @{N='ServerName';E={$Server}}
                        }

                        $entry
                    }
                }
            }
        )

        if (-not $RemoveDuplicate) {
            $result.AddRange($temp)
        }
        else {
            # Check duplicates
            foreach ($o in $temp) {

                if ($skipDupCheck) {
                    $result.Add($o)
                    continue
                }

                # Do a duplicate check based on this property
                if ($o.Distinguishedname) {
                    $dupCheckProp = 'Distinguishedname'
                }
                elseif ($o.Identity) {
                    $dupCheckProp = 'Identity'
                }
                else {
                    Write-Log "Cannot perform duplicate check because the results of '$($Command)' do not have Distinguishedname nor Identity."
                    $skipDupCheck = $true
                    $result.Add($o)
                    continue
                }

                $dups = @($result | Where-Object {$_.$dupCheckProp.ToString() -eq $o.$dupCheckProp.ToString()})

                if ($dups.Count) {
                    Write-Log "`"dropping a duplicate: '$($o.$dupCheckProp.ToString())'`""
                }
                else {
                    $result.Add($o)
                }
            }
        }

        if ($PassThru) {
            Write-Output $result
        }
        else {
            # Extract cmdlet name (e.g "Get-MailboxDatabase" -> "MailboxDatabase")
            $Command.Split(' ')[0] -match ".*-(?<cmdName>.*)" | Out-Null
            $commandName = $Matches['cmdName']
            Save-Object $result -Name $commandName
        }
    }
}

<#
  Write a log to a file and also Write-Verbose
  This automatically creates a file and append
#>
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]$Text,
        [string]$Path = $Script:logPath
    )

    $currentTime = Get-Date
    $currentTimeFormatted = $currentTime.ToString("yyyy/MM/dd HH:mm:ss.fffffff(K)")

    [System.TimeSpan]$delta = 0;
    if ($Script:lastLogTime) {
        $delta = $currentTime.Subtract($Script:lastLogTime)
    }
    else {
        # For the first time, add header
        Add-Content $Path "date-time,delta(ms),info"
    }

    Write-Verbose $Text
    Add-Content $Path "$currentTimeFormatted,$($delta.TotalMilliseconds),$text"
    $Script:lastLogTime = $currentTime
}

<#
  Run Get-*VirtualDirectory & Get-OutlookAnywhere for all servers in $Servers
  If IncludeIISVirtualDirectories is specified, access IIS vdir for Servers == IsDirectAccess.
  Otherwise, only AD info will be collected
#>
function Get-VirtualDirectory {
    [CmdletBinding()]
    param()
    # List of Get-*VirtualDirectory commands.
    # CommantType can be different depending on whether Local PowerShell or Remote PowerShell
    $commands = @(Get-Command Get-*VirtualDirectory -ErrorAction:SilentlyContinue | Where-Object {$_.name -ne 'Get-WebVirtualDirectory' -and $_.name -ne 'Get-VirtualDirectory'})
    $commands += @(Get-Command Get-OutlookAnywhere -ErrorAction:SilentlyContinue)

    foreach ($command in $commands) {
        # If ShowMailboxVirtualDirectories param is available, add it (E2013 & E2016).
        if ($command.Parameters -and $command.Parameters.ContainsKey('ShowMailboxVirtualDirectories')) {
            # if IncludeIISVirtualDirectories, then access direct access servers. otherwise, don't touch servers (only AD)
            if ($IncludeIISVirtualDirectories) {
                Run "$($command.Name) -ShowMailboxVirtualDirectories" -Servers:($allExchangeServers | Where-Object {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -SkipIfNoServers -RemoveDuplicate -PassThru |
                    Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -RemoveDuplicate
            }
            else {
                Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -RemoveDuplicate
            }
        }
        else {
            if ($IncludeIISVirtualDirectories) {
                Run "$($command.Name)" -Servers:($allExchangeServers | Where-Object {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -SkipIfNoServers -RemoveDuplicate -PassThru |
                    Run "$($command.Name) -ADPropertiesOnly" -RemoveDuplicate
            }
            else {
                Run "$($command.Name) -ADPropertiesOnly" -RemoveDuplicate
            }
        }
    }
}

function Invoke-FIPSCmdlet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Servers,
        [Parameter(Mandatory = $true)]
        [string]$FIPSCmdlet
    )

    $result = @(
        foreach($server in $Servers) {
            $command = "Add-PSSnapin -Name Microsoft.Forefront.Filtering.Management.PowerShell;"
            $command += "$FIPSCmdlet;"
            $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($command)

            Write-Log "[$($MyInvocation.MyCommand)] Running $FIPSCmdlet on $server"
            Invoke-Command -ServerName $server -ScriptBlock $scriptblock -ErrorAction SilentlyContinue
        }
    )
    $commandName = $FIPSCmdlet.Substring(4)
    $result | Save-Object -Name $commandName
}

function Invoke-FIPSCommand {
    [CmdletBinding()]
    param(
        [string[]]$Servers
    )

    process {
        # If no Server is given, bail.
        # In PowerShell v2, $null.Count is $null. In V5, $null.Count is 0. Thus, In V2, $null.Count -eq 0 is false, while it's true in v5.
        # To check emptiness, use just $something.Count. If empty, either $null (v2) or 0 (v5), thus it's evaluated to be false in both v2 & v5
        if (-not $Servers.Count) {
            Write-Error ("[$($MyInvocation.MyCommand)] Servers is null or empty")
            return
        }

        $command = "Add-PSSnapin -Name Microsoft.Forefront.Filtering.Management.PowerShell;"
        $command += "Get-Command -Module Microsoft.Forefront.Filtering.Management.PowerShell"
        $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($command)

        # ASSUME all servers have the same FIPS cmdlets
        $FIPSCmdlets = Invoke-Command -ServerName:$Servers[0] -ScriptBlock $scriptblock -ErrorAction SilentlyContinue
        # filter only Get-* cmdlets except Get-ConfigurationValue
        $FIPSCmdlets = $FIPSCmdlets | Where-Object {$_.Name -like "Get-*" -and $_.Name -ne "Get-ConfigurationValue" }

        foreach ($cmdlet in $FIPSCmdlets) {
            Invoke-FIPSCmdlet -Servers:$Servers -FIPSCmdlet:$cmdlet
        }
    }
}

function Get-SPN {
    [CmdletBinding()]
    param (
        # folder path to save output.
        [Parameter(Mandatory = $True)]
        $Path
    )

    # Make sure Path exists; if not, just return error string
    $resolvedPath  = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (-not $resolvedPath) {
        #$PSCmdlet.ThrowTerminatingError((New-Object System.Management.Automation.ErrorRecord "Path '$Path' doesn't exist", $null, ([System.Management.Automation.ErrorCategory]::InvalidData), $null))
        throw "Path '$Path' doesn't exist"
    }

    $filePath = Join-Path -Path $Path -ChildPath "setspn.txt"

    # Check if setspn.exe exists
    $isSetSPNAvailable = $false
    foreach ($path in $env:Path.Split(";")) {
        if ($path) {
            $exePath = Join-Path -Path $Path -ChildPath "setspn.exe"
            if (Test-Path $exePath) {
                $isSetSPNAvailable = $true;
                break;
            }
        }
    }

    if (-not $isSetSPNAvailable) {
        throw "setspn.exe is not available"
    }

    Add-Content -Path:$filePath -Value:"[setspn -P -F -Q http/*]"
    $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q http/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeMDB/*]"
    $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeMDB/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeRFR/*]"
    $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeRFR/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeAB/*]"
    $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeAB/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath
}

function Invoke-ShellCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        $FileName,
        [string]$Argument,
        [switch]$Wait
    )

    $startInfo = New-Object system.diagnostics.ProcessStartInfo
    $startInfo.FileName = $FileName
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = $Argument
    #$startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $startInfo.CreateNoWindow  = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $process.Start() | Out-Null

    if (-not $Wait) {
        Write-Output $process
    }
    else {
        # deadlock can occur b/w parent and child process!
        # https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandardoutput(v=vs.110).aspx

        $stdout = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()

        $result = New-Object -TypeName PSCustomObject -Property @{Process = $process; StdOut = $stdout; ExitCode = $exitCode}
        Write-Output $result
    }
}

function Get-MSInfo32 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        $Servers
    )

    try {
        foreach ($server in $Servers) {
            Write-Log "[$($MyInvocation.MyCommand)] Running on $server"

            $nfoFilePath = Join-Path $Script:Path -ChildPath "$server.nfo"
            $process = Start-Process "msinfo32.exe" -ArgumentList "/Server $server /nfo $nfoFilePath" -PassThru
            if (Get-Process -Id:($process.Id) -ErrorAction:SilentlyContinue) {
                Wait-Process -InputObject:$process
            }
        }
    }
    finally {
        if ($process -and (Get-Process -Id:($process.Id) -ErrorAction:SilentlyContinue)) {
            Write-Error "[$($MyInvocation.MyCommand)] msinfo32 cancelled for $server"
            Stop-Process -InputObject $process
        }
    }
}

function Save-ExchangeEventLog {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        $Path,
        $Server,
        [switch]$IncludeCrimsonLogs,
        [switch]$SkipZip
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
    }

    # By default, collect app and sys logs
    $logs = "Application","System"

    # Save logs from a server into a separate folder
    $savePath = Join-Path $Path -ChildPath $Server
    if (-not (Test-Path $savePath -ErrorAction Stop)) {
        New-Item -ItemType directory $savePath -ErrorAction Stop | Out-Null
    }

    Write-Log "[$($MyInvocation.MyCommand)] Saving event logs on $Server ..."

    # Detect machine-local Window's TEMP path (i.e. C:\Windows\Temp)
    # Logs are saved here temporarily and will be moved to savePath
    $win32os = Get-WmiObject win32_operatingsystem -ComputerName:$Server
    if (-not $win32os) {
        throw "Get-WmiObject win32_operatingsystem failed for '$Server'"
    }

    # This is remote machine's path
    $winTempPath = Join-Path $win32os.WindowsDirectory -ChildPath "Temp"
    $winTempEventPath = Join-Path $winTempPath -ChildPath "EventLogs_$(Get-Date -Format "yyyyMMdd_HHmmss")"
    $uncWinTempEventPath = Join-Path "\\$Server\" -ChildPath $winTempEventPath.Replace(':','$')

    if (-not (Test-Path $uncWinTempEventPath -ErrorAction Stop)) {
        New-Item $uncWinTempEventPath -ItemType Directory -ErrorAction Stop | Out-Null
    }

    # Add crimson logs if requested
    if ($IncludeCrimsonLogs) {
        $logs += (wevtutil el /r:$Server) -like "Microsoft-Exchange*"
    }

    foreach ($log in $logs) {
        # Export event logs to Windows' temp folder
        Write-Log "[$($MyInvocation.MyCommand)] Saving $log ..."
        $fileName = $log.Replace('/', '_') + '.evtx'
        $localFilePath = Join-Path $winTempEventPath -ChildPath $fileName
        wevtutil epl $log $localFilePath /ow /r:$Server
    }

    # Try to zip up before copying in order to save bandwidth unless:
    # - $SkipZip is specified by the caller
    # - Target server is the local machine
    # This is possible only if remote management is enabled on the remote machine (i.e. winrm quickconfig)
    $zipFileName = "EventLogs_$Server.zip"
    $zipCreated = $false

    if (-not $SkipZip -and $env:COMPUTERNAME -ne $Server) {
        try {
            $destination = Join-Path $winTempPath -ChildPath $([Guid]::NewGuid())
            $zipResult = Invoke-Command -ComputerName $Server -ScriptBlock ${function:Compress-Folder} -ArgumentList $winTempEventPath, $destination, $zipFileName -ErrorAction Stop
            $zipCreated = $true
        }
        catch {
            Write-Error "Cannot create a zip file on $Server. Each event log file will be copied. $_"
        }
    }

    if ($zipCreated) {
        Write-Log "[$($MyInvocation.MyCommand)] Copying a zip file '$zipFileName' from $Server"
        $uncZipFilePath = Join-Path "\\$Server\" -ChildPath $zipResult.ZipFilePath.Replace(':','$')
        Move-Item $uncZipFilePath -Destination $savePath -Force
        Remove-Item $([IO.Path]::GetDirectoryName($uncZipFilePath)) -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Log "[$($MyInvocation.MyCommand)] Copying *.evtx files from $Server"
        $evtxFiles = Get-ChildItem -Path $uncWinTempEventPath -Filter '*.evtx'
        foreach ($file in $evtxFiles) {
            Move-Item $file.FullName -Destination $savePath -Force
        }
    }

    # Clean up
    Remove-Item $uncWinTempEventPath -Recurse -Force -ErrorAction SilentlyContinue
}

function Get-DAG {
    [CmdletBinding()]
    param()

    $dags = RunCommand Get-DatabaseAvailabilityGroup
    $result = @(
        foreach ($dag in $dags) {
            # Get-DatabaseAvailabilityGroup with "-Status" fails for cross Exchange versions (e.g. b/w E2010, E2013)
            $dagWithStatus = RunCommand "Get-DatabaseAvailabilityGroup $dag -Status -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue"
            if ($dagWithStatus) {
                $dagWithStatus
            }
            else {
                Write-Log "[$($MyInvocation.MyCommand)] Get-DatabaseAvailabilityGroup $($dag.Name) -Status failed. The result without -Status will be saved."
                $dag
            }
        }
    )

    Save-Object $result -Name "DatabaseAvailabilityGroup"
}

function Get-DotNetVersion {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$Server = $env:COMPUTERNAME
    )

    begin {}

    process {
        # Read NDP registry
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
        }
        catch {
            $msg = $_.Exception.Message.Replace("`r`n", "")
            throw "Couldn't open registry key of $Server. $msg"
        }

        $ndpKey = $reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP")
        $result = @(
            foreach ($versionKeyName in $ndpKey.GetSubKeyNames())  {
                # ignore "CDF" etc
                if ($versionKeyName -notlike "v*") {
                    continue
                }

                $versionKey = $ndpKey.OpenSubKey($versionKeyName)
                $version = $versionKey.GetValue("Version", "")
                $sp = $versionKey.GetValue("SP", "")
                $install = $versionKey.GetValue("Install", "")

                if ($version) {
                    New-Object PSCustomObject -Property @{Version = $version; SP = $sp; Install = $install; SubKey = $null; Release = $release; NET45Version = $null; ServerName = $Server}
                    continue
                }

                # for v4 and V4.0, check sub keys
                foreach ($subKeyName in $versionKey.GetSubKeyNames()) {

                    $subKey = $versionKey.OpenSubKey($subKeyName)
                    $version = $subKey.GetValue("Version", "")
                    $install = $subKey.GetValue("Install", "")
                    $release = $subKey.GetValue("Release", "")

                    if ($release) {
                        $NET45Version = Get-Net45Version $release
                    }
                    else {
                        $NET45Version = $null
                    }

                    New-Object PSCustomObject -Property @{Version = $version; SP = $sp; Install = $install; SubKey = $subKeyName;Release = $release; NET45Version = $NET45Version; ServerName = $Server}
                }
            }
        )

        $result = $result | Sort-Object -Property Version
        Write-Output $result
    } # end of process{}

    end {}
}

function Get-Net45Version {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory=$True)]
        $Release
    )

    switch ($Release) {
        {$_ -ge 528040} {'4.8 or later'; break}
        {$_ -ge 461808} {'4.7.2'; break}
        {$_ -ge 460798} {'4.7'; break}
        {$_ -ge 394802} {"4.6.2"; break}
        {$_ -ge 394254} {"4.6.1"; break}
        {$_ -ge 393295} {"4.6"; break}
        {$_ -ge 379893} {"4.5.2"; break}
        {$_ -ge 378675} {'4.5.1'; break}
        {$_ -ge 378389} {'4.5'; break}
        default {$null}
    }
}

function Get-TlsRegistry {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline=$true)]
        [string]$Server= $env:COMPUTERNAME
    )

    Begin{}

    Process {
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
        }
        catch {
            Write-Error "Couldn't open registry key of $Server.`n$_"
            return
        }

        $result = New-Object System.Collections.Generic.List[object]

        # OS SChannel related
        $protocols = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\')
        # "Protocols" key should exist
        foreach ($protocolKeyName in $protocols.GetSubKeyNames()) {
            # subKeyName is "SSL 2.0", "TLS 1.0", etc
            $protocolKey = $protocols.OpenSubKey($protocolKeyName)
            #$subkeyNames = @('Client','Server')
            foreach ($subKeyName in $protocolKey.GetSubKeyNames()) {
                $subKey = $protocolKey.OpenSubKey($subKeyName)

                $disabledByDefault = $subKey.GetValue('DisabledByDefault', '')
                $enabled = $subKey.GetValue('Enabled', '')

                $result.Add((New-Object PSCustomObject -Property @{ServerName = $Server
                    Name = "SChannel $protocolKeyName $subKeyName"
                    DisabledByDefault = $disabledByDefault
                    Enabled = $enabled
                    RegistryKey = $subKey.Name
                    })
                )
            }
        }

        # .NET related
        $netKeyNames = @('SOFTWARE\Microsoft\.NETFramework\', 'SOFTWARE\Wow6432Node\Microsoft\.NETFramework\')

        foreach ($netKeyName in $netKeyNames) {
            $netKey = $reg.OpenSubKey($netKeyName)
            $netSubKeyNames = @('v2.0.50727','v4.0.30319')

            foreach ($subKeyName in $netSubKeyNames) {
                $subKey = $netKey.OpenSubKey($subKeyName)
                if (-not $subKey) {
                    continue
                }

                $systemDefaultTlsVersions = $subKey.GetValue('SystemDefaultTlsVersions','')
                $schUseStrongCrypto = $subKey.GetValue('SchUseStrongCrypto','')

                if ($subKey.Name.IndexOf('Wow6432Node', [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $name = ".NET Framework $subKeyName (Wow6432Node)"
                }
                else {
                    $name = ".NET Framework $subKeyName"
                }

                $result.Add((New-Object PSCustomObject -Property @{ServerName = $Server
                    Name = $name
                    SystemDefaultTlsVersions = $systemDefaultTlsVersions
                    SchUseStrongCrypto = $schUseStrongCrypto
                    RegistryKey = $subKey.Name
                    })
                )
            }
        }

        $result
    }

    End{}
}

function Get-TCPIP6Registry {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline=$true)]
        [string]$Server= $env:COMPUTERNAME
    )

    begin{}

    process {
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
        }
        catch {
            Write-Error "Couldn't open registry key of $Server.`n$_"
            return
        }

        $key = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\')
        $disabledComponents = $key.GetValue('DisabledComponents','')

        New-Object PSCustomObject -Propert @{DisabledComponents = $disabledComponents}
    }

    end{}
}

function Get-IISWebBinding {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline=$true)]
        [string]$Server
    )

    begin {
        $block = {
            $err = Import-Module WebAdministration 2>&1
            if ($err) {
                Write-Error "Import-Module WebAdministration failed."
                return
            }

            Get-WebBinding
        }
    }

    process {
        if ($Server -eq $env:COMPUTERNAME) {
            Invoke-Command -ScriptBlock $block
            return
        }

        try {
            $sess = New-PSSession -ComputerName $Server -ErrorAction Stop
            Invoke-Command -Session $sess -ScriptBlock $block
        }
        catch {
            Write-Error "Failed to invoke command on a remote session to $Server.$_"
        }
        finally {
            if ($sess) {
                Remove-PSSession $sess
            }
        }
    }
    end{}
}

function Save-PerfmonLog {
    [CmdletBinding()]
    param(
        $Path,
        $Server,
        [switch]$SkipZip
    )

    if (-not (Test-Path $Path)){
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    # Save logs from a Server into a separate folder
    $savePath = Join-Path $Path -ChildPath $Server
    if (-not (Test-Path $savePath -ErrorAction Stop)) {
        New-Item -ItemType directory $savePath -ErrorAction Stop | Out-Null
    }

    # Find the perfmon path on the remote machine
    $win32os = Get-WmiObject win32_operatingsystem -ComputerName $Server
    $win32env = Get-WmiObject win32_environment -ComputerName $Server
    if (-not $win32os -or -not $win32env) {
        # WMI failed. Maybe wrong Server name?
        throw "Get-WmiObject win32_operatingsystem or win32_environment failed for '$Server'"
    }

    # This is remote machine's paths
    $exchangePath = $win32env | Where-Object {$_.Name -eq 'ExchangeInstallPath'}
    if ($exchangePath) {
        $perfmonPath = Join-Path $exchangePath.VariableValue "Logging\Diagnostics\DailyPerformanceLogs"
    }
    else {
        throw "Cannt find ExchangeInstallPath on $Server"
    }

    # Try to compress before copying unless:
    # - $SkipZip is specified by the caller
    # - Target server is the local machine
    $zipCreated = $false
    if (-not $SkipZip -and $env:COMPUTERNAME -ne $Server) {
        # Compess the perfmon logs & save it to Windows's TEMP path.
        $winTempPath = Join-Path $win32os.WindowsDirectory -ChildPath "Temp"
        $winTempPerfmonPath = Join-Path $winTempPath -ChildPath "Perfmon_$(Get-Date -Format "yyyyMMdd_HHmmss")"
        $uncWinTempPerfmonPath = "\\$Server\" + $winTempPerfmonPath.Replace(':','$')

        $zipFileName = "Perfmon_$Server.zip"

        try {
            Write-Progress -Activity "Compressing perfmon logs on $Server" -Status "Started (This might take a while)" -PercentComplete -1
            $zipResult = Invoke-Command -ComputerName $Server -ScriptBlock ${function:Compress-Folder} -ArgumentList $perfmonPath,$winTempPerfmonPath,$zipFileName -ErrorAction Stop
            Write-Progress -Activity "Compressing perfmon logs on $Server" -Status "Done" -Completed
            $zipCreated = $true
        }
        catch {
            Write-Error "Cannot create a zip file on $Server. Each event log file will be copied. $_"
        }
    }

    if ($zipCreated) {
        Write-Progress -Activity "Copying a perfmon zip file from $Server" -Status "Started (This might take a while)" -PercentComplete -1
        $uncZipFilePath = Join-Path "\\$Server\" -ChildPath $zipResult.ZipFilePath.Replace(':','$')
        Move-Item $uncZipFilePath -Destination $savePath
        Write-Progress -Activity "Copying a perfmon zip file from $Server" -Status "Done" -Completed
    }
    else {
        # Manually copy perfmon logs
        $uncPerfmonPath = Join-Path "\\$Server\" -ChildPath $perfmonPath.Replace(':', '$')
        $count = 1
        $files = @(Get-ChildItem -Path "$uncPerfmonPath\*" -Include "*.blg")
        foreach ($file in $files) {
            Write-Progress -Activity "Copying perfmon logs from $Server" -Status "$count/$($files.Count)" -PercentComplete $($count/$files.Count*100)
            Copy-Item $file -Destination $savePath
            $count++
        }
    }

    if ($uncWinTempPerfmonPath) {
        Remove-Item $uncWinTempPerfmonPath -Force -Recurse -ErrorAction SilentlyContinue
    }
}


<#
  Main
#>

# If the path doesn't exist, create it.
if (-not (Test-Path $Path -ErrorAction Stop)) {
    New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
}
$Path = Resolve-Path $Path

$cmd = Get-Command "Get-OrganizationConfig" -ErrorAction:SilentlyContinue
if (-not $cmd) {
    throw "Get-OrganizationConfig is not available. Please run with Exchange Remote PowerShell session"
}
$OrgConfig = Get-OrganizationConfig
$OrgName = $orgConfig.Name
$IsExchangeOnline = $orgConfig.LegacyExchangeDN.StartsWith('/o=ExchangeLabs')


# Create a temporary folder to store data
$tempFolder = New-Item $(Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())) -ItemType directory -ErrorAction Stop

# Prepare for logging
# NOTE: until $logPath is defined, don't call Write-Log
$logFileName = "Log.txt"
$logPath = Join-Path -Path $tempFolder.FullName -ChildPath $logFileName

$lastLogTime = $null
Write-Log "Organization Name = $OrgName"
Write-Log "Script Version = $version"
Write-Log "COMPUTERNAME = $env:COMPUTERNAME"
Write-Log "IsExchangeOnline = $IsExchangeOnline"

# Log parameters (raw values are in $PSBoundParameters, but want fixed-up values (e.g. Path)
$sb = New-Object System.Text.StringBuilder
foreach ($paramName in $PSBoundParameters.Keys) {
    $var = Get-Variable $paramName -ErrorAction SilentlyContinue
    if ($var) {
        $sb.Append("$($var.Name):$($var.Value); ") | Out-Null
    }
}
Write-Log $sb.ToString()

# Switch Path to the temporary folder so that all the items will be saved there
$originalPath = $Path
$Path = $tempFolder.FullName
Write-Log "Temporary Folder: $($tempFolder.FullName)"


if ($Servers -and -not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    throw "Servers parameter is specified, but Get-ExchangeServer is not available."
}

# Prepare the list of Exchange Servers to directly access by parsing the values specified in "Servers" parameter
# Used in VDir, Mailbox Catabase Copy, Certificate etc.
# First, get the candidates from the user specified values in $Servers
$directAccessCandidates =@(
    foreach ($server in $Servers) {
        # $Server's value might be something like "e2013*" and matches multiple Servers
        $exServers = @(Get-ExchangeServer $server -ErrorAction SilentlyContinue)

        if (-not $exServers.Count) {
            Write-Log "Get-ExchangeServer did not find any Server matching '$Server'"
        }

        foreach ($exServer in $exServers) {
            # In PowerShellv2. $exServer may be $null.
            if (-not $exServer) {
                continue
            }

            # Skip Edge servers unless it's the local server.
            if ($exServer.IsEdgeServer -and $env:COMPUTERNAME -ne $exServer.Name) {
                Write-Log "Dropping $($exServer.Name) from directAccessCandidates since it's an Edge server"
                continue
            }

            # add if it's not a duplicate
            $inDAS = @($directAccessCandidates | Where-Object {$_.Name -eq $exServer.Name}).Count -gt 0
            if (-not $inDAS) {
                $exServer
            }
        }
    }
)

Write-Log "directAccessCandidates = $directAccessCandidates"

# Now test connectivity to those servers
# Since there shouldn't be anything blocking communication b/w Exchange Servers, we should be able to use ICMP
$directAccessServers = @(
    foreach ($server in $directAccessCandidates) {
        if (Test-Connection -ComputerName:$server.Name -Count 1 -Quiet) {
            $server
        }
        else {
            Write-Log "Connectivity test failed on $server"
        }
    }
)
Write-Log "directAccessServers = $directAccessServers"

if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
    $allExchangeServers = @(Get-ExchangeServer)
    $allExchangeServers | Add-Member -Type NoteProperty -Name IsDirectAccess -Value:$false
}
else {
    $allExchangeServers = @()
}

foreach ($server in $allExchangeServers) {
    if (@($directAccessServers | Where-Object {$_.Name -eq $server}).Count -gt 0) {
        $server.IsDirectAccess = $true
    }
}

# Save errors for troubleshooting purpose
# $errs = New-Object System.Collections.Generic.List[object]

#
# Start collecting
#
$transcriptPath = Join-Path -Path $Path -ChildPath "transcript.txt"
$transcriptEnabled = $false
try {
    Start-Transcript -Path $transcriptPath -NoClobber -ErrorAction:Stop
    $transcriptEnabled = $true
}
catch {
    Write-Log "Start-Transcript is not available"
}

# Start of try for transcript
try {
# Write-Progress's Activity string
$collectionActivity = "Collecting Data"

# org settings
Write-Progress -Activity:$collectionActivity -Status:"Org Settings" -PercentComplete:0

# When you don't specify 'Path' for Save-Object, it's saved to $Script:Path
Run Get-OrganizationConfig
Run Get-AdminAuditLogConfig
Run Get-AvailabilityAddressSpace
Run Get-AvailabilityConfig
Run Get-OrganizationRelationship
Run Get-ADServerSettings
Run Get-AuthConfig
Run Get-AuthRedirect
Run Get-AuthServer
Run Get-DomainController
Run Get-IRMConfiguration
Run Get-OfflineAddressBook
Run Get-OrganizationalUnit
Run Get-OutlookProvider
Run Get-OwaMailboxPolicy
Run Get-ResourceConfig
Run Get-SmimeConfig
Run Get-UserPrincipalNamesSuffix
Write-Log "Org done"

# ActiveSync
Write-Progress -Activity:$collectionActivity -Status:"ActiveSync Settings" -PercentComplete:10
Run Get-ActiveSyncDeviceAccessRule
Run Get-ActiveSyncDeviceAutoblockThreshold
Run Get-ActiveSyncDeviceClass
Run "Get-ActiveSyncMailboxPolicy -WarningAction:SilentlyContinue"
Run Get-MobileDeviceMailboxPolicy
Run Get-ActiveSyncOrganizationSettings

# Transport Settings
Write-Progress -Activity:$collectionActivity -Status:"Transport Settings" -PercentComplete:20
Run Get-TransportConfig
Run Get-AcceptedDomain
Run Get-ReceiveConnector
Run Get-SendConnector
Run Get-ForeignConnector
Run Get-RemoteDomain
Run Get-ClassificationRuleCollection
Run Get-ContentFilterConfig
Run Get-ContentFilterPhrase
#Run Get-DataClassification
Run Get-DeliveryAgentConnector
Run Get-DlpPolicy
# Run Get-DlpPolicyTemplate
Run Get-EdgeSubscription
Run Get-EdgeSyncServiceConfig
Run Get-EmailAddressPolicy
Run Get-HostedContentFilterRule
Run Get-IPAllowListConfig
Run Get-IPAllowListEntry -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsHubTransportServer}) -SkipIfNoServers
Run Get-IPAllowListProvider
Run Get-IPAllowListProvidersConfig
Run Get-IPBlockListConfig
Run Get-IPBlockListEntry -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsHubTransportServer}) -SkipIfNoServers
Run Get-IPBlockListProvider
Run Get-IPBlockListProvidersConfig
Run Get-JournalRule
Run Get-RecipientFilterConfig
Run Get-RMSTemplate
Run Get-SenderFilterConfig
Run Get-SenderIdConfig
Run Get-SenderReputationConfig
Run Get-TransportRule
# these cmdlets are meant to run locally and don't have Server specifiers (-Server, -Identity)
#Run Get-TransportAgent
#Run Get-TransportPipeline

Write-Log "Transport done"

# AD Setting
Write-Progress -Activity:$collectionActivity -Status:"AD Settings" -PercentComplete:30
Run Get-ADSite
Run Get-AdSiteLink

Run Get-ExchangeAssistanceConfig

# AddressBook
Run Get-GlobalAddressList
Run Get-AddressList
Run Get-AddressBookPolicy

# Retention
Run Get-RetentionPolicy
Run Get-RetentionPolicyTag
Write-Log "AD AddressBook Retention Done"

# WMI
Write-Progress -Activity $collectionActivity -Status:"Server Settings" -PercentComplete:40

Run Get-ExchangeServer
Run Get-MailboxServer

# For CAS (>= E14) in DAS list, include ASA info
Run "Get-ClientAccessServer -IncludeAlternateServiceAccountCredentialStatus -WarningAction:SilentlyContinue" -Servers:($allExchangeServers | Where-Object {$_.IsDirectAccess -and $_.IsClientAccessServer -and -$_.IsE14OrLater}) -Identifier:Identity -SkipIfNoServers -RemoveDuplicate -PassThru |
    Run "Get-ClientAccessServer -WarningAction:SilentlyContinue" -Identifier:Identity -RemoveDuplicate

Run Get-ClientAccessArray
Run Get-RpcClientAccess
Run "Get-TransportServer -WarningAction:SilentlyContinue"
Run Get-TransportService
Run Get-FrontendTransportService
Run Get-ExchangeDiagnosticInfo
Run Get-ExchangeServerAccessLicense

Run Get-PopSettings -Servers:$allExchangeServers
Run Get-ImapSettings -Servers:$allExchangeServers

Write-Log "Server Done"

# Database
Write-Progress -Activity $collectionActivity -Status:"Database Settings" -PercentComplete:50

Run "Get-MailboxDatabase -Status -IncludePreExchange" -Servers:($allExchangeServers | Where-Object {$_.IsMailboxServer -and $_.IsDirectAccess}) -SkipIfNoServers -RemoveDuplicate -PassThru |
    Run "Get-MailboxDatabase -IncludePreExchange" -RemoveDuplicate

Run "Get-PublicFolderDatabase -Status" -Servers:($allExchangeServers | Where-Object {$_.IsMailboxServer -and $_.IsDirectAccess}) -SkipIfNoServers -RemoveDuplicate -PassThru |
    Run "Get-PublicFolderDatabase" -RemoveDuplicate

Run Get-MailboxDatabaseCopyStatus -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsMailboxServer}) -SkipIfNoServers
Get-DAG
Run Get-DatabaseAvailabilityGroupConfiguration
if (Get-Command Get-DatabaseAvailabilityGroup -ErrorAction:SilentlyContinue) {
    Run "Get-DatabaseAvailabilityGroupNetwork -ErrorAction:SilentlyContinue" -Servers:(Get-DatabaseAvailabilityGroup) -Identifier:'Identity'
}
Write-Log "Database Done"

# Virtual Directories
Write-Progress -Activity $collectionActivity -Status:"Virtual Directory Settings" -PercentComplete:60
Run 'Get-VirtualDirectory'
Run "Get-IISWebBinding" -Servers $directAccessServers -SkipIfNoServers -PassThru | Save-Object -Name WebBinding

# Active Monitoring & Managed Availability
Write-Progress -Activity $collectionActivity -Status:"Monitoring Settings" -PercentComplete:70
Run Get-GlobalMonitoringOverride
Run Get-ServerMonitoringOverride -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -SkipIfNoServers
Run Get-ServerComponentState -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity -SkipIfNoServers
Run Get-HealthReport -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity -SkipIfNoServers
Run Get-ServerHealth -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity -SkipIfNoServers
Run Test-ServiceHealth -Servers:$directAccessServers -SkipIfNoServers

# Federation & Hybrid
Write-Progress -Activity $collectionActivity -Status:"Monitoring Settings" -PercentComplete:75
Run Get-SharingPolicy
Run Get-HybridConfiguration
Run Get-FederationTrust
Run Get-FederatedOrganizationIdentifier
#Run Get-FederationInformation
#Run Get-FederatedDomainProof
Run "Get-IntraOrganizationConfiguration -WarningAction:SilentlyContinue"
Run Get-IntraOrganizationConnector
Run Get-InboundConnector
Run Get-OutboundConnector

# Exchange Certificate
Write-Progress -Activity $collectionActivity -Status:"Exchange Certificate" -PercentComplete:80
Run Get-ExchangeCertificate -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater}) -SkipIfNoServers

# Throttling
Write-Progress -Activity $collectionActivity -Status:"Throttling" -PercentComplete:85
Run Get-ThrottlingPolicy
Run Get-ThrottlingPolicyAssociation

# misc
Write-Progress -Activity $collectionActivity -Status:"Misc" -PercentComplete:85
Run Get-MigrationConfig
Run Get-MigrationEndpoint
Run Get-NetworkConnectionInfo -Servers:$directAccessServers -Identifier:Identity -SkipIfNoServers
Run Get-ProcessInfo -Servers:$directAccessServers -Identifier:TargetMachine -SkipIfNoServers
Run Get-OutlookProtectionRule
Run Get-PolicyTipConfig
Run Get-RbacDiagnosticInfo
Run Get-RoleAssignmentPolicy
Run Get-SearchDocumentFormat
Run Get-MailboxAuditBypassAssociation
Run Get-SettingOverride
Run "Get-Mailbox -Arbitration" -PassThru | Save-Object -Name 'Mailbox-Arbitration'
Run "Get-Mailbox -Monitoring" -PassThru | Save-Object -Name 'Mailbox-Monitoring'
Run "Get-Mailbox -PublicFolder" -PassThru | Save-Object -Name 'Mailbox-PublicFolder'
Run Get-UMService

# FIPS
Run Get-MalwareFilteringServer
Run Get-MalwareFilterPolicy
Run Get-MalwareFilterRule
if ($IncludeFIPS) {
    Write-Progress -Activity $collectionActivity -Status:"FIPS" -PercentComplete:85
    Run Invoke-FIPSCommand -Servers ($directAccessServers | Where-Object {$_.IsE15OrLater -and $_.IsHubTransportServer}) -SkipIfNoServers
}

Run Get-HostedConnectionFilterPolicy
Run Get-HostedContentFilterPolicy
Run Get-HostedContentFilterRule
Run Get-AntiPhishPolicy
Run Get-AntiPhishRule
Run "Get-PhishFilterPolicy -SpoofAllowBlockList -Detailed"

# .NET Framework Versions
Run Get-DotNetVersion -Servers:($directAccessServers) -Identifier:Server -SkipIfNoServers

# TLS Settings
Run Get-TlsRegistry -Servers $directAccessServers -Identifier:Server -SkipIfNoServers

# TCPIP6
Run Get-TCPIP6Registry -Servers $directAccessServers -Identifier:Server -SkipIfNoServers

# MSInfo32
# Get-MSInfo32 -Servers $directAccessServers

# WMI
# Win32_powerplan is available in Win7 & above.
Run 'Get-WmiObject -namespace root\cimv2\power -class Win32_PowerPlan' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_PowerPlan
Run 'Get-WmiObject -Class Win32_PageFileSetting' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_PageFileSetting
Run 'Get-WmiObject -Class Win32_ComputerSystem' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_ComputerSystem
Run 'Get-WmiObject -Class Win32_OperatingSystem' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_OperatingSystem
Run "Get-WmiObject -Class Win32_NetworkAdapterConfiguration" -Servers:$directAccessServers -Identifier:ComputerName -SkipIfNoServers -PassThru |
    Where-Object {$_.IPEnabled} | Save-Object -Name Win32_NetworkAdapterConfiguration

#Run "Get-WmiObject Win32_Process" -Servers:$directAccessServers -Identifier:ComputerName -SkipIfNoServers

if ($IsExchangeOnline) {
    Write-Log "Skipping Get-SPN & Invoke-Ldifde since this is an Exchange Online Organization"
}
else {
    Run "Get-SPN -Path:$Path"

    # Ldife for Exchange Org
    Write-Progress -Activity $collectionActivity -Status:"Running Ldifde" -PercentComplete:90
    Run "Invoke-Ldifde -Path:$Path"
}

# Collect EventLogs
if ($IncludeEventLogs -or $IncludeEventLogsWithCrimson) {
    Write-Progress -Activity $collectionActivity -Status:"Event Logs" -PercentComplete:90

    $eventLogPath = Join-Path $Path -ChildPath 'EventLogs'
    if ($IncludeEventLogsWithCrimson) {
        # Get-ExchangeEventLog -Path:$eventLogPath -Servers:$directAccessServers -IncludeCrimsonLogs
        Run "Save-ExchangeEventLog -Path:$eventLogPath -IncludeCrimsonLogs" -Servers $directAccessServers
    }
    else {
        # Get-ExchangeEventLog -Path:$eventLogPath -Servers:$directAccessServers
        Run "Save-ExchangeEventLog -Path $eventLogPath" -Servers $directAccessServers
    }
}

# Collect Perfmon Log
if ($IncludePerformanceLog) {
    Write-Progress -Activity $collectionActivity -Status:"Perfmon Logs" -PercentComplete:90
    Run "Save-PerfmonLog -Path:$(Join-Path $Path 'Perfmon')" -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater}) -SkipIfNoServers
}

# Save errors
if ($Script:errs.Count) {
    $errPath = Join-Path $Path -ChildPath "Error"
    if (-not (Test-Path errPath)) {
        New-Item $errPath -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $errs | Export-Clixml $(Join-Path $errPath "errs.xml") -Depth 5
}

} # end of try for transcript
finally {
    # release transcript file even when script is stopped in the middle.
    if ($transcriptEnabled) {
        Stop-Transcript
    }
}

Write-Progress -Activity $collectionActivity -Status:"Packing into a zip file" -PercentComplete:95
Write-Log "Running Compress-Folder -Path:$Path -ZipFileName:$OrgName -RemoveFiles:(-not $KeepOutputFiles) -Destination:$originalPath"
Compress-Folder -Path:$Path -ZipFileName:$OrgName -RemoveFiles:(-not $KeepOutputFiles) -Destination:$originalPath -IncludeDateTime | Out-Null

if (-not $KeepOutputFiles){
    $err = $(Remove-Item $Path -Force) 2>&1
    if ($err) {
        Write-Warning "Failed to delete a temporary folder `"$Path`""
    }
}
else {
    Write-Warning "Temporary folder `"$Path`" contains files collected"
}

Write-Progress -Activity $collectionActivity -Status "Done" -Completed
Write-Output "Done!"