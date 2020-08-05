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
Switch to include FIPS (Forefront Information Protection Service) related information from the servers specified in "Servers" parameter.

.PARAMETER IncludeEventLogs
Switch to include Application & System event logs from the servers specified in "Servers" parameter.

.PARAMETER IncludeEventLogsWithCrimson
Switch to include Exchange-related Crimson logs ("Microsoft-Exchange-*") as well as Application & System event logs from the servers specified in "Servers" parameter.

.PARAMETER IncludePerformanceLog
Switch to include Exchange's Perfmon log from the servers specified in "Servers" parameter (Only Exchange 2013 and above collects perfmon log by default).

.PARAMETER IncludeIISLog
Switch to include IIS log from the servers specified in "Servers" parameter.

.PARAMETER IncludeExchangeLog
List of log folders unders %ExchangeInstallPath%Logging to collect from the servers specified in "Servers" parameter.

.PARAMETER IncludeTransportLog
List of transport-related log folders to collect from the servers specified in "Servers" parameter (e.g. Connectivity, MessagingTracing etc)

.PARAMETER IncludeFastSearchLog
Switch to include FAST Search log from the servers specified in "Servers" parameter

.PARAMETER FromDateTime
Log files whose LastWriteTime is greater than or equal to this value are collected for the following log types:
IncludePerformanceLog, IncludeIISLog, IncludeExchangeLog, IncludeTransportLog, and IncludeFastSearchLog.

.PARAMETER ToDateTime
Log files whose LastWriteTime is less than or equal to this value are collected for the following log types:
IncludePerformanceLog, IncludeIISLog, IncludeExchangeLog, IncludeTransportLog, and IncludeFastSearchLog.

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
See https://github.com/jpmessaging/CollectExchangeInfo

Copyright 2020 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
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
    [switch]$IncludeIISLog,
    [string[]]$IncludeExchangeLog,
    [ValidateSet('Connectivity', 'MessageTracking','SendProtocol', 'ReceiveProtocol', 'RoutingTable', 'Queue')]
    [string[]]$IncludeTransportLog,
    [switch]$IncludeFastSearchLog,
    [Nullable[DateTime]]$FromDateTime,
    [Nullable[DateTime]]$ToDateTime,
    [switch]$KeepOutputFiles
)

$version = "2020-08-06"
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

function Compress-Folder {
    [CmdletBinding()]
    param(
        # Specifies a path to one or more locations.
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$Destination,
        [string]$ZipFileName,
        [Nullable[DateTime]]$FromDateTime,
        [Nullable[DateTime]]$ToDateTime,
        [switch]$IncludeDateTime,
        [switch]$RemoveFiles,
        [switch]$UseShellApplication
    )

    if (-not (Test-Path $Path)) {
        throw [System.IO.DirectoryNotFoundException]"Path '$Path' is not found"
    }

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

    if (Test-Path $zipFilePath) {
        # Append a randome string to the zip file name.
        $zipFileName = $zipFileNameWithouExt + "_" + [System.IO.Path]::GetRandomFileName().Substring(0,8) + '.zip'
        $zipFilePath = Join-Path $Destination -ChildPath $zipFileName
    }

    $NETFileSystemAvailable = $false

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        # Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $NETFileSystemAvailable = $true
    }
    catch {
        Write-Warning "System.IO.Compression.FileSystem wasn't found. Using alternate method"
    }

    $files = @(Get-ChildItem $Path -Recurse | Where-Object {-not $_.PSIsContainer})

    # Apply filters
    if ($FromDateTime)  {
        $files = $files | Where-Object {$_.LastWriteTime -ge $FromDateTime}
    }

    if ($ToDateTime) {
        $files = $files | Where-Object {$_.LastWriteTime -le $ToDateTime}
    }

    # If there's no files, bail.
    if ($files.Count -eq 0) {
        New-Object PSCustomObject -Property @{
            ZipFilePath = $null
            FilesRemoved = $false
        }
        return
    }

    if ($NETFileSystemAvailable -and $UseShellApplication -eq $false) {
        # Note: [System.IO.Compression.ZipFile]::CreateFromDirectory() fails when one or more files in the directory is locked.
        #[System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $zipFilePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

        $zipStream = $zipArchive = $null
        try {
            New-Item $zipFilePath -ItemType file | Out-Null

            $zipStream = New-Object System.IO.FileStream -ArgumentList $zipFilePath, ([IO.FileMode]::Open)
            $zipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipStream, ([IO.Compression.ZipArchiveMode]::Create)
            $count = 0

            foreach ($file in $files) {
                Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Adding $($file.FullName)" -PercentComplete (100 * $count / $files.Count)

                $fileStream = $zipEntryStream = $null
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

            Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Done" -Completed
        }
    }
    else {
        # Use Shell.Application COM

        # Create a zip file manually
        $shellApp = New-Object -ComObject Shell.Application
        Set-Content $zipFilePath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (Get-Item $zipFilePath).IsReadOnly = $false
        $zipFile = $shellApp.NameSpace($zipFilePath)

        # Add the entire folder.
        if ($null -eq $FromDateTime -and $null -eq $ToDateTime ) {
            # Start copying the whole and wait until it's done. CopyHere works asynchronously.
            $zipFile.CopyHere($Path)

            # Now wait and poll
            $inProgress = $true
            $delayMilliseconds = 200
            Start-Sleep -Milliseconds 3000
            [System.IO.FileStream]$file = $null
            while ($inProgress) {
                Start-Sleep -Milliseconds $delayMilliseconds
                $file = $null
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
        else {
            # Copy the files to a temporary folder (with folder hierarchy), and then zip it.
            $tempPath = Join-Path $(Get-WindowsTempFolder -Server $env:COMPUTERNAME) -ChildPath "$([Guid]::NewGuid().ToString())\$(Split-Path $Path -Leaf)"
            New-Item $tempPath -ItemType directory -ErrorAction stop | Out-Null

            try {
                foreach ($fileInfo in $files) {
                    $tempDest = $tempPath
                    if ($fileInfo.DirectoryName.Length -gt $Path.Length) {
                        $folderName = $fileInfo.DirectoryName.Substring($Path.Length + 1)
                        $tempDest = Join-Path $tempDest -ChildPath $folderName
                        if (-not (Test-Path $tempDest)) {
                            New-Item $tempDest -ItemType Directory | Out-Null
                        }
                    }
                    Copy-Item -Path $fileInfo.FullName -Destination $tempDest
                }

                $zipFile.CopyHere($tempPath)

                # Now wait and poll
                $inProgress = $true
                $delaymsec = 20
                $maxDelaymsec = 200

                Start-Sleep -Milliseconds 200

                while ($inProgress) {
                    Start-Sleep -Milliseconds $delaymsec
                    $file = $null
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
                        else {
                            $delaymsec = $delaymsec * 2
                            if ($delaymsec -ge $maxDelaymsec) {
                                $delaymsec = $maxDelaymsec
                            }
                        }
                    }
                }
            }
            finally {
                # Remove the temporary folder
                if (Test-Path $tempPath) {
                    Remove-Item (Get-Item $tempPath).Parent.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        } # end of else

        Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Done"  -Completed
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

# Convert a local path to UNC path.
# C:\temp --> \\myServer\C$\temp
# These functions are meant to be just small helper and not bullet-proof.
function ConvertTo-UNCPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        $Server
    )

    "\\$Server\$($Path.Replace(':', '$'))"
}

function ConvertFrom-UNCPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Given Path can be a local path or remote path: e.g.
    # "\\Server\C$\temp\etc\" or "C:\temp\etc"
    if (-not ($Path -match  '(\\\\(?<Server>[^\\]+)\\)?(?<Path>.*)')) {
        throw "$Path looks invalid (or bug here)"
    }

    $server = $Matches['Server']
    $localPath = $Matches['Path'].Replace('$', ':')
    New-Object PSCustomObject -Property @{
        Server = $server
        LocalPath = $localPath
    }
}

function Save-Item {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourcePath, # Path to save
        [Parameter(Mandatory = $true)]
        $DestitionPath, # Where to save
        $Filter = '*', # filter works only when copied without zipping first.
        $ZipFileName,
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime,
        [switch]$SkipZip # Switch to disable make a zip file before copying.
    )

    if (-not (Test-Path $SourcePath)) {
        throw "$SourcePath is not found"
    }

    if (-not (Test-Path $DestitionPath)) {
        New-Item $DestitionPath -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $DestitionPath = Resolve-Path $DestitionPath

    $serverAndPath = ConvertFrom-UNCPath $SourcePath
    $server = $serverAndPath.Server
    $localPath = $serverAndPath.LocalPath

    # Remember the servers that cannot be zip remotely and skip zipping next time.
    if (-not $script:SkipZipServers) {
        $script:SkipZipServers = @{}
    }

    if ($script:SkipZipServers.ContainsKey($Server)) {
        $SkipZip = $true
        Write-Log "[$($MyInvocation.MyCommand)] Skip zipping for $Server"
    }

    # Try to zip up remotely before copying.
    $zipCreated = $false
    if ($server -and $env:COMPUTERNAME -ne $server -and -not $SkipZip) {
        # Compress & save it to Windows's TEMP path.

        # Form the zip file name
        if (-not $ZipFileName) {
            $ZipFileName = $SourcePath.Substring($SourcePath.LastIndexOf('\') + 1)
        }
        if (-not ($ZipFileName.EndsWith(".zip"))) {
            $zipFileName = "$ZipFileName.zip"
        }

        Write-Progress -Activity "Compressing $localPath on $server" -Status "Started (This might take a while)" -PercentComplete -1
        try {
            $winTempPath = Get-WindowsTempFolder -Server $server
            #$zipResult = Compress-Folder -Path $localPath -Destination $winTempPath -ZipFileName $zipFileName -ErrorAction Stop
            $zipResult = Invoke-Command -ComputerName $server -ScriptBlock ${function:Compress-Folder} -ArgumentList $localPath,$winTempPath,$zipFileName,$FromDateTime,$ToDateTime -ErrorAction Stop
            $zipCreated = ($null -ne $zipResult.ZipFilePath)
        }
        catch {
            Write-Error "Cannot create a zip file on $Server. Each log file will be copied. $_"
            $script:SkipZipServers.Add($Server,$null)
        }

        Write-Progress -Activity "Compressing $localPath on $Server" -Status "Done" -Completed

    }

    if ($zipCreated) {
        Write-Progress -Activity "Copying a zip file from $server" -Status "Started (This might take a while)" -PercentComplete -1
        $uncZipFile = ConvertTo-UNCPath $zipResult.ZipFilePath -Server $server
        Move-Item $uncZipFile -Destination $DestitionPath
        Write-Progress -Activity  "Copying a zip file from $server" -Status "Done" -Completed
    }
    else {
        # Manually copy
        #Copy-Item $SourcePath\* -Destination $DestitionPath -Recurse -Filter $Filter -Force
        $files = @(Get-ChildItem $SourcePath -Recurse | Where-Object {-not $_.PSIsContainer})

        if ($FromDateTime)  {
            $files = $files | Where-Object {$_.LastWriteTime -ge $FromDateTime}
        }

        if ($ToDateTime) {
            $files = $files | Where-Object {$_.LastWriteTime -le $ToDateTime}
        }

        foreach ($file in $files) {
            $destination = Join-Path $DestitionPath $file.Directory.Name
            if (-not (Test-Path $destination)) {
                New-Item $destination -ItemType Directory | Out-Null
            }
            Copy-Item $file.FullName -Destination $destination -Force
        }

        if ($files.Count -eq 0) {
            Write-Log "[$($MyInvocation.MyCommand)] There're no files in $SourcePath from $FromDateTime to $ToDateTime"
        }
    }
}

function Save-IISLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        $Server,
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime
    )

    # Find the locations of IIS logs on the Server (can be multiple locations)
    $webSites = @(
        try {
            $session = $null
            $session = New-PSSession -ComputerName $Server -ErrorAction Stop
            Invoke-Command -Session $session -ScriptBlock {
                Import-Module WebAdministration
                $webSites = @(Get-Website)
                foreach ($webSite in $webSites) {
                    # The directory might contain environment variable (e.g. %SystemDrive%\inetpub\logs\LogFiles).
                    $directory = [System.Environment]::ExpandEnvironmentVariables($webSite.logFile.directory)
                    New-Object PSCustomObject -Property @{
                        SiteName = $webSite.Name
                        Directory = $directory
                    }
                }
            }
            $webSiteFound = $true
        }
        catch {
            # ignored
        }
        finally {
            if ($session) {
                Remove-PSSession $session
            }
        }
    )

    if ($webSiteFound) {
        foreach ($webSiteGroup in $($webSites | Group-Object Directory)) {
            # Form a folder name.
            # There can be multiple web sites with different log directories. Save each directory to a different locations
            $folderName = $null
            foreach ($site in $webSiteGroup.Group) {
                $folderName += $site.SiteName + '&'
            }
            $folderName = $folderName.Remove($folderName.Length - 1)
            $destination = Join-Path $Path -ChildPath "$Server\$folderName"

            $uncPath = ConvertTo-UNCPath $webSiteGroup.Group[0].Directory -Server $Server
            Save-Item -SourcePath $uncPath -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
        }
    }
    else {
        # Web sites information is not found (maybe PSSession cannot be established)
        # Try the default IIS log location ('c:\inetpub\logs\LogFiles')
        $uncPath = ConvertTo-UNCPath 'C:\inetpub\logs\LogFiles' -Server $Server
        if (Test-Path $uncPath) {
            $destination = Join-Path $Path -ChildPath $Server
            Save-Item -SourcePath $uncPath -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
        }
        else {
            # Give up
            Write-Error "Cannot find the IIS log directory of server $Server and also cannot find $uncPath"
        }
    }
}

function Save-HttpErr {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        $Server,
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime
    )

    # The path of HTTPERR log can be changed by:
    # HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\HTTP\Parameters\ErrorLoggingDir
    # But this should be rare. So just assume all servers use the default path.
    $logPath = Join-Path $env:SystemRoot 'System32\LogFiles\HTTPERR'

    $source = ConvertTo-UNCPath $logPath -Server $Server
    $destination = Join-Path $Path -ChildPath $Server
    Save-Item -SourcePath $source -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
}

<#
Save folder under %ExchangeInstallPath%Logging
#>
function Save-ExchangeLogging {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path, # destination
        [Parameter(Mandatory=$true)]
        $Server,
        [Parameter(Mandatory=$true)]
        $FolderPath, # subfolder path under %ExchangeInstallPath%Logging
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime
    )

    $logPath = $null

    # Diagnostics path can be modified. So update the folder path if necessary
    if ($FolderPath -like 'Diagnostics\*') {
        $customPath = $null
        try {
            $customPath = Get-DiagnosticsPath -Server $Server -ErrorAction SilentlyContinue
        }
        catch {
            Write-Error "Get-DiagnosticsPath failed. $_."
        }

        if ($customPath) {
            $subPath = $FolderPath.Substring($FolderPath.IndexOf('\') + 1)
            $logPath = Join-Path $customPath -ChildPath $subPath
            Write-Log "[$($MyInvocation.MyCommand)] Custom Diagnostics path is found. Using $logPath"
        }
    }

    # Default path: %ExchangeInstallPath% + $FolderPath
    if (-not $logPath) {
        $exchangePath  = Get-ExchangeInstallPath -Server $Server
        $logPath = Join-Path $exchangePath "Logging\$FolderPath"
    }

    $source = ConvertTo-UNCPath $logPath -Server $Server
    $destination = Join-path $Path -ChildPath $Server
    Save-Item -SourcePath $source -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
}

function Save-TransportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $Server,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Connectivity', 'MessageTracking','SendProtocol', 'ReceiveProtocol', 'RoutingTable', 'Queue')]
        $Type,
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime
    )

    $transport = $null
    if (Get-Command 'Get-TransportService' -ErrorAction SilentlyContinue) {
        $transport = Get-TransportService $Server
    }
    elseif (Get-Command 'Get-TransportServer' -ErrorAction SilentlyContinue) {
        $transport = Get-TransportServer $Server
    }

    # If both Get-TransportService & Get-TransportServer are not available, bail.
    if (-not $transport) {
        throw "Get-TransportService/TransportServer is not available."
    }

    $frontendTransport = $null
    if (Get-Command 'Get-FrontendTransportService' -ErrorAction SilentlyContinue) {
        $frontendTransport = Get-FrontendTransportService $Server
    }

    foreach ($logType in $Type) {
        # Parameter name is ***LogPath
        $paramName = $logType + 'LogPath'
        if (-not $transport.$paramName) {
            Write-Error "Cannot find $paramName in the result of Get-TransportService"
            continue
        }
        $sourcePath = ConvertTo-UNCPath $transport.$paramName.ToString() -Server $Server
        $destination = Join-path $Path -ChildPath "$Server\Hub"
        Save-Item -SourcePath $sourcePath -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime

        if ($frontendTransport -and $frontendTransport.$paramName) {
            $sourcePath = ConvertTo-UNCPath $frontendTransport.$paramName.ToString() -Server $Server
            $destination = Join-path $Path -ChildPath "$Server\FrontEnd"
            Save-Item -SourcePath $sourcePath -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
        }
    }
}

function Get-DiagnosticsPath {
    [CmdletBinding()]
    param($Server)

    $reg = $diagKey = $path = $null
    try {
        # Get the value of "HKEY_LOCAL_MACHINE\Software\Microsoft\ExchangeServer\v15\Diagnostics\LogFolderPath" if exits.
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
        $diagKey = $reg.OpenSubKey("Software\Microsoft\ExchangeServer\v15\Diagnostics\")
        if (-not $diagKey) {
            throw "OpenSubKey failed for 'Software\Microsoft\ExchangeServer\v15\Diagnostics\' on $Server"
        }
        $path = $diagKey.GetValue('LogFolderPath')
    }
    finally {
        if ($diagKey) { $diagKey.Close() }
        if ($reg) { $reg.Close() }
    }

    return $path
}

function Save-ExchangeSetupLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $Server
    )

    $source = ConvertTo-UNCPath 'C:\ExchangeSetupLogs' -Server $Server
    $destination = Join-path $Path -ChildPath $Server
    Save-Item -SourcePath $source -DestitionPath $destination
}

function Save-FastSearchLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $Server,
        [DateTime]$FromDateTime,
        [DateTime]$ToDateTime
    )

    $exsetupPath = Get-ExchangeInstallPath -Server $Server
    $source = ConvertTo-UNCPath $(Join-Path $exsetupPath 'Bin\Search\Ceres\Diagnostics\Logs') -Server $Server
    $destination = Join-path $Path -ChildPath $Server
    Save-Item -SourcePath $source -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime
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

    if ($Script:OrgConfig) {
        $exorg = $Script:OrgConfig.DistinguishedName
    }
    else {
        $exorg = (Get-OrganizationConfig).DistinguishedName
    }

    if (-not $exorg) {
        throw "Couldn't get Exchange org DN"
    }

    # If this is an Edge server, use a port 50389.
    $port = 0
    $server = Get-ExchangeServer $env:COMPUTERNAME -ErrorAction SilentlyContinue
    if ($server -and $server.IsEdgeServer) {
        $port = 50389
    }

    $fileNameWihtoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $stdOutput = Join-Path $resolvedPath -ChildPath "$fileNameWihtoutExtension.out"

    if ($port) {
        $result = Invoke-ShellCommand -FileName 'ldifde' -Argument "-d `"$exorg`" -s localhost -t $port -f `"$filePath`""
    }
    else {
        $result = Invoke-ShellCommand -FileName 'ldifde' -Argument "-d `"$exorg`" -f `"$filePath`""
    }

    $result.StdOut | Out-File $stdOutput -Encoding utf8

    if ($result.ExitCode -ne 0) {
        throw "ldifde failed. exit code = $($result.ExitCode)."
    }
}

<#
Get an available runspace. If available runspace is not found, it creates a new one (local or remote runspace)
#>
function Get-Runspace {
    [CmdletBinding()]
    param()

    # For the first time, create a runspace pool
    if ($null -eq $Script:RunspacePool) {
        $Script:RunspacePool = New-Object System.Collections.Generic.List[System.Management.Automation.Runspaces.Runspace]
    }

    # For the first time, determine local or remote runspace
    if (-not $Script:ExchangeLocalPS -and -not $Script:ExchangeRemotePS) {
        $command = Get-Command "Get-OrganizationConfig"

        if ($Command.CommandType -eq [System.Management.Automation.CommandTypes]::Cmdlet -and $Command.ModuleName -eq 'Microsoft.Exchange.Management.PowerShell.E2010') {
            $Script:ExchangeLocalPS = $true
        }
        elseif ($Command.CommandType -eq [System.Management.Automation.CommandTypes]::Function -and $Command.Module) {
            $Script:ExchangeRemotePS = $true

            # Remember the primary runspace so that its ConnectionInfo can be used when creating a new remote runspace.
            $Script:PrimaryRunspace = Get-PSSession | Where-Object {$_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.Availability -eq [System.Management.Automation.Runspaces.RunspaceAvailability]::Available -and  $_.Runspace.ConnectionInfo.ConnectionUri.ToString() -notlike '*ps.compliance.protection.outlook.com*'} | Select-Object -First 1 -ExpandProperty Runspace
            if (-not $Script:PrimaryRunspace) {
                # If "Available" runspace is not there, then select whichever
                $Script:PrimaryRunspace = Get-PSSession | Where-Object {$_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.Runspace.ConnectionInfo.ConnectionUri.ToString() -notlike '*ps.compliance.protection.outlook.com*'} | Select-Object -First 1 -ExpandProperty Runspace
            }

            $Script:RunspacePool.Add($Script:PrimaryRunspace)
        }
    }

    # Find an available runspace
    $rs = $Script:RunspacePool | Where-Object {$_.RunspaceAvailability -eq [System.Management.Automation.Runspaces.RunspaceAvailability]::Available} | Select-Object -First 1
    if ($rs) {
        return $rs
    }

    # If there's no available runspace, create one.
    if ($Script:ExchangeLocalPS) {
        $rs = [RunspaceFactory]::CreateRunspace()
        $rs.Open()

        # Add Exchange Local PowerShell so that it's ready to be used.
        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        $ps.AddCommand('Add-PSSnapin').AddParameter('Name','Microsoft.Exchange.Management.PowerShell.E2010') | Out-Null
        $ps.Invoke() | Out-Null
        $ps.Dispose()
    }
    elseif ($Script:ExchangeRemotePS) {
        $rs = [RunspaceFactory]::CreateRunspace($Script:PrimaryRunspace.ConnectionInfo)
        $rs.Open()
    }

    Write-Log "$(if ($rs.ConnectionInfo) {'Remote'} else {'Local'}) runspace was created. Runspace count: $($Script:RunspacePool.Count + 1)"
    $Script:RunspacePool.Add($rs)
    Write-Output $rs
}

function Remove-Runspace {
    [CmdletBinding()]
    param()

    $count = 0
    foreach ($rs in $Script:RunspacePool) {
        if ($rs -ne $Script:PrimaryRunspace) {
            $rs.Dispose()
            ++$count
        }
    }

    Write-Log "$count runspaces were removed"
}

<#
Helper function to create an AsyncCallback instance which invokes the given scriptblock callback.
Basically same as:
https://web.archive.org/web/20190222052659/http://www.nivot.org/blog/post/2009/10/09/PowerShell20AsynchronousCallbacksFromNET
#>
function New-AsyncCallback {
    param (
        [parameter(Mandatory=$true)]
        [scriptblock]$Callback
    )

    # Class that exposes an event of type AsyncCallback that Register-ObjectEvent can register to.
    $AsyncCallbackProxyType = @"
        using System;
        using System.Threading;

        public sealed class AsyncCallbackProxy
        {
            // This is the exposed event. The sole purpose is for Register-ObjectEvent to hook to.
            public event AsyncCallback AsyncOpComplete;

            // Private ctor
            private AsyncCallbackProxy() { }

            // Raise the event
            private void OnAsyncOpComplete(IAsyncResult ar)
            {
                // For .NET 2.0, System.Threading.Volatile.Read is not available.
                //AsyncCallback temp = System.Threading.Volatile.Read(ref AsyncOpComplete);
                AsyncCallback temp = AsyncOpComplete;
                if (temp != null) {
                    temp(ar);
                }
            }

            // This is the AsyncCallback instance.
            public AsyncCallback Callback
            {
                get { return new AsyncCallback(OnAsyncOpComplete); }
            }

            public static AsyncCallbackProxy Create()
            {
                return new AsyncCallbackProxy();
            }
        }
"@

    if (-not ("AsyncCallbackProxy" -as [type])) {
        Add-Type $AsyncCallbackProxyType
    }

    $proxy = [AsyncCallbackProxy]::Create()
    Register-ObjectEvent -InputObject $proxy -EventName AsyncOpComplete -Action $Callback -Messagedata $args | Out-Null

    # When an async operation finishes, this AsyncCallback instance gets invoked, which in turn raises AsynOpCompleted event of the proxy object.
    # Since this AsynOpCompleted is registered by Register-ObjectEvent, it calls the script block.
    $proxy.Callback
}

<#
  Run a given command only if it's available
  Run with parameters specified as Global Parameter (i.e. $script:Parameters)
#>
function RunCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        [int]$TimeoutSeconds
    )

    $endOfCmdlet = $Command.IndexOf(" ")
    if ($endOfCmdlet -lt 0) {
        $cmdlet = $Command
    }
    else {
        $cmdlet = $Command.Substring(0, $endOfCmdlet)
    }

    # Check if cmdlet is available
    $cmd = Get-Command $cmdlet -ErrorAction:SilentlyContinue
    if ($null -eq $cmd) {
        Write-Log "$cmdlet is not available"
        return
    }

    $ExchangeLocalPS = $false
    $ExchangeRemotePS = $false
    if ($cmd.CommandType -eq [System.Management.Automation.CommandTypes]::Cmdlet -and $cmd.ModuleName -eq 'Microsoft.Exchange.Management.PowerShell.E2010') {
        $ExchangeLocalPS = $true
    }
    elseif ($cmd.CommandType -eq [System.Management.Automation.CommandTypes]::Function -and $cmd.Module -and $cmd.Module.ExportedFunctions.ContainsKey("Get-OrganizationConfig")) {
        $ExchangeRemotePS = $true
    }

    [System.Management.Automation.PowerShell]$ps = $null
    if ($ExchangeLocalPS -or $ExchangeRemotePS) {
        $ps = [PowerShell]::Create()
        $ps.Runspace = Get-Runspace
    }

    if ($ps) {
        $psCommand = $ps.AddCommand($cmdlet, $true)
    }

    # Check parameters.
    # If any explicitly-requested parameter is not available, ignore.
    $paramMatches = Select-String "(\s-(?<paramName>\w+))((\s+|:)\s*(?<paramVal>[^-]\S+))?" -Input $Command -AllMatches

    if ($paramMatches) {
        $paramList = @(
        foreach($paramMatch in $paramMatches.Matches) {
            $paramName = $paramMatch.Groups['paramName'].Value
            $paramValue = $paramMatch.Groups['paramVal'].Value

            $params = @(
            foreach ($param in $cmd.Parameters.GetEnumerator()) {
                if ($param.Key -like "$paramName*") {
                    $param
                }
            })

            # If there's no match or too many matches, ignore.
            if ($params.Count -eq 0) {
                Write-Log "Parameter '$paramName' is not available for $cmdlet"
                continue
            }
            elseif ($params.Count -gt 1) {
                Write-Log "Parameter '$paramName' is ambiguous for $cmdlet"
                continue
            }

            if ($ps -and $params[0].Value.SwitchParameter) {
                $psCommand.AddParameter($params[0].Key) | Out-Null
            }
            elseif ($ps) {
                $psCommand.AddParameter($params[0].Key, $paramValue) | Out-Null
            }

            Write-Output $params[0]
        }
        ) # end of $paramList array subexpression
    }

    # Check if any parameter is requested globally. Ignore the parameter if it's not available for this cmdlet.
    foreach ($param in $script:Parameters) {
        $paramName = ($param -split ":")[0]
        $paramVal = ($param -split ":")[1]
        if ($cmd.Parameters[$paramName]) {
            # Explicitly-requested parameters take precedence; If not already in the list, add it.
            if ($paramList.Key -notcontains $paramName) {
                $Command += " -$param"
                if ($ps) {
                    if ($paramVal) {
                        $psCommand.AddParameter($paramName, $paramVal) | Out-Null
                    }
                    else {
                        $psCommand.AddParameter($paramName) | Out-Null
                    }

                }
           }
        }
    }

    # Finally run the command
    Write-Log "Running $Command $(if ($ps -and $TimeoutSeconds){"with $TimeoutSeconds seconds timeout."})"

    $timeoutmsec = -1
    if ($TimeoutSeconds) {
        $timeoutmsec = $TimeoutSeconds * 1000
    }

    $ar = $errs = $o = $null
    try {
        if ($ps) {
            $ar = $ps.BeginInvoke()
            if ($ar.AsyncWaitHandle.WaitOne($timeoutmsec)) {
                $o = $ps.EndInvoke($ar)
                $errs = @($ps.Streams.Error)
            }
            else {
                Write-Log "[Timeout] '$Command' timed out after $TimeoutSeconds seconds"
            }
        }
        else {
            $errs = @($($o = Invoke-Expression $Command) 2>&1)
        }
    }
    catch {
        # Log the terminating error.
        Write-Log "[Terminating Error] '$Command' failed. $($_.ToString()) $(if ($_.Exception.Line) {"(At line:$($_.Exception.Line) char:$($_.Exception.Offset))"})"
        if ($null -ne $Script:errs) {$Script.errs.Add($_)}
    }
    finally {
        if ($errs.Count) {
            foreach ($err in $errs) {
                Write-Log "[Non-Terminating Error] Error in '$Command'. $($err.ToString()) $(if ($err.Exception.Line) {"(At line:$($err.Exception.Line) char:$($err.Exception.Offset))"})"
            }
        }

        if ($null -ne $o) {
            Write-Output $o
        }

        if ($ps) {
            if ($ps.InvocationStateInfo.State -eq "Running") {
                # Asychronously stop the command and dispose the powershell instance.
                $context = New-Object PSCustomObject -Property @{
                    PowerShell = $ps
                    AsyncResult = $ar
                }

                $ps.BeginStop(
                    (
                        New-AsyncCallback {
                            param ([IAsyncResult]$asyncResult)
                            $state = $asyncResult.AsyncState
                            $state.PowerShell.Dispose()
                            $state.AsyncResult.AsyncWaitHandle.Close()
                        }
                    ),
                    $context
                ) | Out-Null
            }
            else {
                $ps.Dispose()
                $ar.AsyncWaitHandle.Close()
            }
        }
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
        [switch]$PassThru,
        [int]$TimeoutSeconds = 180
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
        # Check if cmdlet is available. If not, bail (RunCommand check the availability of cmdlet. So this is just an optimization)
        $endOfCmdlet = $Command.IndexOf(" ")
        if ($endOfCmdlet -lt 0) {
            $cmdlet = $Command
        }
        else {
            $cmdlet = $Command.Substring(0, $endOfCmdlet)
        }

        $cmd = Get-Command $cmdlet -ErrorAction:SilentlyContinue
        if (-not $cmd) {
            Write-Log "$cmdlet is not available"
            return
        }

        $temp = @(
            if (-not $Servers.Count -and -not $SkipIfNoServers) {
                RunCommand $Command -TimeoutSeconds $TimeoutSeconds
            }
            elseif ($Servers) {
                foreach ($Server in $Servers) {
                    $firstTimeAddingServerName = $true
                    foreach ($entry in @(RunCommand "$Command -$Identifier $Server" -TimeoutSeconds $TimeoutSeconds)) {
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

        # Deserialize if SerializationData property is available.
        if (-not $Script:formatter) {
            $Script:formatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        }

        for ($i = 0; $i -lt $temp.Count; ++$i) {
            if ($null -ne $temp[$i].serializationData) {
                try {
                    $stream = New-Object system.io.memoryStream -ArgumentList (, $temp[$i].serializationData)
                    $temp[$i] = $Script:formatter.Deserialize($stream)
                }
                catch {
                    Write-Log "Deserialization failed. $($_.ToString())"
                }
                finally {
                    if ($stream) {
                        $stream.Dispose()
                    }
                }
            }
        }

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
                    Write-Log "Cannot perform duplicate check because the results of '$($Command)' have neither Distinguishedname nor Identity."
                    $skipDupCheck = $true
                    $result.Add($o)
                    continue
                }

                $dups = @($result | Where-Object {$_.$dupCheckProp.ToString() -eq $o.$dupCheckProp.ToString()})

                if ($dups.Count) {
                    Write-Log "Dropping a duplicate: '$($o.$dupCheckProp.ToString())'"
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
  Write a log to a file.  This automatically creates a file and append to it.
  Make sure to call Close-Log so that data in buffer is flushed and release the file handle.
#>
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]$Text,
        [string]$Path = $Script:logPath
    )

    $currentTime = Get-Date
    $currentTimeFormatted = $currentTime.ToString('o')

    if (-not $Script:logWriter) {
        # For the first time, open file & add header
        [IO.StreamWriter]$Script:logWriter = [IO.File]::AppendText($Path)
        $Script:logWriter.WriteLine("date-time,delta(ms),info")
    }

    [TimeSpan]$delta = 0;
    if ($Script:lastLogTime) {
        $delta = $currentTime.Subtract($Script:lastLogTime)
    }

    # Format as CSV:
    $sb = New-Object System.Text.StringBuilder
    $sb.Append($currentTimeFormatted).Append(',') | Out-Null
    $sb.Append($delta.TotalMilliseconds).Append(',') | Out-Null
    $sb.Append('"').Append($Text.Replace('"', "'")).Append('"') | Out-Null

    $Script:logWriter.WriteLine($sb.ToString())

    $sb = $null
    $Script:lastLogTime = $currentTime
}

function Close-Log {
    if ($Script:logWriter) {
        $Script:logWriter.Close()
        $Script:logWriter = $null
    }
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

function Invoke-FIPS {
    [CmdletBinding()]
    param(
    [string[]]$Servers
    )

    if (-not $Servers.Count) {
        return
    }

    # key: Cmdlet Name, value: List of cmdlet output
    $resultSet = @{}

    foreach ($server in $Servers) {
        $session = $null
        $FIPSCmdlets = $null
        try {
            # First, setup a session/runspace; load PSSnapin and obtain FIPS-related cmdlets
            $command = "Add-PSSnapin -Name Microsoft.Forefront.Filtering.Management.PowerShell;"
            $command += "Get-Command -Module Microsoft.Forefront.Filtering.Management.PowerShell"
            $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($command)

            $session = New-PSSession -ComputerName $server
            $FIPSCmdlets = @(Invoke-Command -Session $session -ScriptBlock $scriptblock -ErrorAction SilentlyContinue `
                | Where-Object {$_.Name -like "Get-*" -and $_.Name -ne "Get-ConfigurationValue"})

            foreach ($cmdlet in $FIPSCmdlets){
                $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($cmdlet)
                Write-Log "Running $cmdlet on $server"

                try {
                    $errs = @($($o = Invoke-Command -Session $session -ScriptBlock $scriptblock) 2>&1)

                    if ($errs.Count) {
                        foreach ($err in $errs) {
                            Write-Log "[Non-Terminiating Error]$err"
                        }
                    }

                    $cmdletName = $cmdlet.ToString().Substring(4)
                    if ($null -eq $resultSet[$cmdletName]) {
                        $resultSet[$cmdletName] = New-Object System.Collections.Generic.List[object]
                    }

                    $resultSet[$cmdletName].Add($o)
                }
                catch {
                    Write-Log "[Terminating Error] $_"
                }
            }
        }
        catch {
            Write-Log "[Terminating Error] Failed to setup PSSession. $_"
        }
        finally {
            $session | Remove-PSSession -ErrorAction SilentlyContinue
        }
    }

    # Save results
    foreach ($result in $resultSet.GetEnumerator()) {
        $result.Value | Save-Object -Name $result.Key
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

    $writer = $null
    try {
        $writer = [IO.File]::AppendText($filePath)
        $writer.WriteLine("[setspn -P -F -Q http/*]")
        $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q http/*'
        $writer.WriteLine($result.StdOut)

        $writer.WriteLine("[setspn -P -F -Q exchangeMDB/*]")
        $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeMDB/*'
        $writer.WriteLine($result.StdOut)

        $writer.WriteLine("[setspn -P -F -Q exchangeRFR/*]")
        $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeRFR/*'
        $writer.WriteLine($result.StdOut)

        $writer.WriteLine("[setspn -P -F -Q exchangeAB/*]")
        $result = Invoke-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeAB/*'
        $writer.WriteLine($result.StdOut)
    }
    finally {
        if ($writer) {
            $writer.Close()
        }
    }
}

function Invoke-ShellCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        $FileName,
        [string]$Argument
    )

    $startInfo = New-Object system.diagnostics.ProcessStartInfo
    $startInfo.FileName = $FileName
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = $Argument
    #$startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $startInfo.CreateNoWindow  = $true

    $process = $job = $null

    try {
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo

        $stdErr = New-Object System.Text.StringBuilder
        $job = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -MessageData $stdErr -Action {
            $stdErr = $event.MessageData
            $stdErr.Append($eventArgs.Data)
        }

        # stdOut can be read asynchronously, but there's no need. thus commented out.
        # $stdOut = New-Object System.Text.StringBuilder

        # $jobStdOut = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -MessageData $stdOut -Action {
        #     $stdOut = $event.MessageData
        #     $stdOut.Append($eventArgs.Data)
        # }

        $process.Start() | Out-Null

        # Be careful here. Deadlock can occur b/w parent and child process!
        # https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandardoutput(v=vs.110).aspx
        $processId = $process.Id
        $process.BeginErrorReadLine()

        #$process.BeginOutputReadLine()
        $stdOut = $process.StandardOutput.ReadToEnd()

        $process.WaitForExit()
        $exitCode = $process.ExitCode

        New-Object -TypeName PSCustomObject -Property @{PID = $processId; StdOut = $stdOut; StdErr = $stdErr.ToString(); ExitCode = $exitCode}

    }
    finally {
        if ($job) {
            Stop-Job $job
            Remove-Job $job
            $job.Dispose()
        }

        if ($process) {
            if (-not $process.HasExited) {
                Stop-Process -InputObject $process -Force
            }
            $process.Dispose()
        }
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

    # Save logs from a server into a separate folder
    $destination = Join-Path $Path -ChildPath $Server
    if (-not (Test-Path $destination)) {
        New-Item -ItemType directory $destination -ErrorAction Stop | Out-Null
    }

    # This is remote machine's path
    $winTempPath = Get-WindowsTempFolder -Server $Server
    $winTempEventPath = Join-Path $winTempPath -ChildPath "EventLogs_$(Get-Date -Format "yyyyMMdd_HHmmss")"
    $uncWinTempEventPath = ConvertTo-UNCPath $winTempEventPath -Server $Server

    if (-not (Test-Path $uncWinTempEventPath -ErrorAction Stop)) {
        New-Item $uncWinTempEventPath -ItemType Directory -ErrorAction Stop | Out-Null
    }

    Write-Log "[$($MyInvocation.MyCommand)] Saving event logs on $Server ..."
    # By default, collect app and sys logs
    $logs = "Application","System"

    # Add crimson logs if requested
    if ($IncludeCrimsonLogs) {
        $logs += (wevtutil el /r:$Server) -like 'Microsoft-Exchange*'

        # This is for the FAST Search.
        $logs += (wevtutil el /r:$Server) -like 'Microsoft-Office Server*'
    }

    foreach ($log in $logs) {
        # Export event logs to Windows' temp folder
        Write-Log "[$($MyInvocation.MyCommand)] Saving $log ..."
        $fileName = $log.Replace('/', '_') + '.evtx'
        $localFilePath = Join-Path $winTempEventPath -ChildPath $fileName
        wevtutil epl $log $localFilePath /ow /r:$Server
    }

    Save-Item -SourcePath $uncWinTempEventPath -DestitionPath $destination
    Remove-Item $uncWinTempEventPath -Recurse -Force -ErrorAction SilentlyContinue
}

<#
Return the Windows' TEMP folder for a given server.
This function will throw on failure.
#>
function Get-WindowsTempFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Server
    )

    # Cache the result of WMI
    if ($null -eq $Script:Win32OSCache) {
        $Script:Win32OSCache = @{}
    }

    if ($Script:Win32OSCache.ContainsKey($Server)) {
        $win32os = $Script:Win32OSCache[$Server]
    }
    else {
        $win32os = Get-WmiObject win32_operatingsystem -ComputerName $Server
        if (-not $win32os) {
            throw "Get-WmiObject win32_operatingsystem failed for '$Server'"
        }
        $Script:Win32OSCache.Add($Server, $win32os)
    }

    Join-Path $win32os.WindowsDirectory -ChildPath "Temp"
}

<#
Return the value of ExchangeInstallPath environment variable for a given server.
This function will throw on failure.
#>
function Get-ExchangeInstallPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Server
    )

    # Cache the result of WMI win32_environment
    if ($null -eq $Script:Win32EnvCache) {
        $Script:Win32EnvCache = @{}
    }

    if ($Script:Win32EnvCache.ContainsKey($Server)) {
        $win32env = $Script:Win32EnvCache[$Server]
    }
    else {
        $win32env = Get-WmiObject win32_environment -ComputerName $Server
        $Script:Win32EnvCache.Add($Server, $win32env)
    }

    $exchangePath = $win32env | Where-Object {$_.Name -eq 'ExchangeInstallPath'}
    if (-not $exchangePath) {
        throw "Cannt find ExchangeInstallPath on $Server"
    }

    $exchangePath.VariableValue
}

function Get-DAG {
    [CmdletBinding()]
    param()

    $dags = RunCommand Get-DatabaseAvailabilityGroup

    if (-not $dags.Count) {
        return
    }

    $result = @(
        foreach ($dag in $dags) {
            # Get-DatabaseAvailabilityGroup with "-Status" fails for cross Exchange versions (e.g. b/w E2010, E2013)
            # This could take a long time before it fails. Add a timeout.
            $dagWithStatus = RunCommand "Get-DatabaseAvailabilityGroup $dag -Status -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue" -TimeoutSeconds 180
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
        $reg = $ndpKey = $null
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
            $ndpKey = $reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP")
            if (-not $ndpKey) {
                throw "OpenSubKey failed on 'SOFTWARE\Microsoft\NET Framework Setup\NDP'."
            }

            $result = @(
                foreach ($versionKeyName in $ndpKey.GetSubKeyNames())  {
                    $versionKey = $null
                    try {
                        # ignore "CDF" etc
                        if ($versionKeyName -notlike "v*") {
                            continue
                        }

                        $versionKey = $ndpKey.OpenSubKey($versionKeyName)
                        if (-not $versionKey) {
                            Write-Error "OpenSubKey failed on $versionKeyName. Skipping."
                            continue
                        }

                        $version = $versionKey.GetValue("Version", "")
                        $sp = $versionKey.GetValue("SP", "")
                        $install = $versionKey.GetValue("Install", "")

                        if ($version) {
                            New-Object PSCustomObject -Property @{
                                Version = $version
                                SP = $sp
                                Install = $install
                                SubKey = $null
                                Release = $release
                                NET45Version = $null
                                ServerName = $Server
                            }

                            continue
                        }

                        # for v4 and V4.0, check sub keys
                        foreach ($subKeyName in $versionKey.GetSubKeyNames()) {
                            $subKey = $null
                            try {
                                $subKey = $versionKey.OpenSubKey($subKeyName)
                                if (-not $subKey) {
                                    Write-Error "OpenSubKey failed on $subKeyName. Skipping."
                                    continue
                                }

                                $version = $subKey.GetValue("Version", "")
                                $install = $subKey.GetValue("Install", "")
                                $release = $subKey.GetValue("Release", "")

                                if ($release) {
                                    $NET45Version = Get-Net45Version $release
                                }
                                else {
                                    $NET45Version = $null
                                }

                                New-Object PSCustomObject -Property @{
                                    Version = $version
                                    SP = $sp
                                    Install = $install
                                    SubKey = $subKeyName
                                    Release = $release
                                    NET45Version = $NET45Version
                                    ServerName = $Server
                                }
                            }
                            finally {
                                if ($subKey) {$subKey.Close()}
                            }
                        }
                    }
                    finally {
                        if ($versionKey) {$versionKey.Close()}
                    }
                }
            )

            $result = $result | Sort-Object -Property Version
            Write-Output $result
        }
        finally {
            if ($ndpKey) { $ndpKey.Close() }
            if ($reg) { $reg.Close() }
        }

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
    $reg = $protocols = $null
    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)

        # OS SChannel related
        $protocols = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\')

        # Note: OpenSubKey returns $null if the operation failed.
        if ($protocols) {
            foreach ($protocolKeyName in $protocols.GetSubKeyNames()) {
                $protocolKey = $null
                try {
                    # subKeyName is "SSL 2.0", "TLS 1.0", etc
                    $protocolKey = $protocols.OpenSubKey($protocolKeyName)
                    if (-not $protocolKey) {
                        Write-Error "OpenSubKey failed for $protocolKeyName on $Server. Skipping."
                        continue
                    }

                    foreach ($subKeyName in $protocolKey.GetSubKeyNames()) {
                        $subKey = $null
                        try {
                            $subKey = $protocolKey.OpenSubKey($subKeyName)
                            if (-not $subKey) {
                                Write-Error "OpenSubKey failed for $subKeyName on $Server. Skipping."
                                continue
                            }

                            $disabledByDefault = $subKey.GetValue('DisabledByDefault', '')
                            $enabled = $subKey.GetValue('Enabled', '')

                            New-Object PSCustomObject -Property @{
                                ServerName = $Server
                                Name = "SChannel $protocolKeyName $subKeyName"
                                DisabledByDefault = $disabledByDefault
                                Enabled = $enabled
                                RegistryKey = $subKey.Name
                            }
                        }
                        finally {
                            if ($subKey) {$subKey.Close()}
                        }
                    }
                }
                finally {
                    if ($protocolKey) {$protocolKey.Close()}
                }
            }
        }
        else {
            # If OpenSubKey failed, write to error stream and flow through.
            Write-Error "OpenSubKey failed for 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\' on $Server"
        }

        # .NET related
        $netKeyNames = @('SOFTWARE\Microsoft\.NETFramework\', 'SOFTWARE\Wow6432Node\Microsoft\.NETFramework\')
        foreach ($netKeyName in $netKeyNames) {
            $netKey = $null
            try {
                $netKey = $reg.OpenSubKey($netKeyName)
                if (-not $netKey) {
                    Write-Error "OpenSubKey failed on $netKeyName on $Server. Skipping."
                    continue
                }

                $netSubKeyNames = @('v2.0.50727','v4.0.30319')

                foreach ($subKeyName in $netSubKeyNames) {
                    $subKey = $null
                    try {
                        $subKey = $netKey.OpenSubKey($subKeyName)
                        if (-not $subKey) {
                            Write-Error "OpenSubKey failed for $subKeyName on $Server. Skipping."
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

                        New-Object PSCustomObject -Property @{
                            ServerName = $Server
                            Name = $name
                            SystemDefaultTlsVersions = $systemDefaultTlsVersions
                            SchUseStrongCrypto = $schUseStrongCrypto
                            RegistryKey = $subKey.Name
                        }
                    }
                    finally {
                        if ($subKey) {$subKey.Close()}
                    }
                }
            }
            finally {
                if ($netKey) { $netKey.Close() }
            }
        }
    }
    finally {
        if ($protocols) { $protocols.Close() }
        if ($reg) { $reg.Close() }
    }

    } # End of process{}

    End{}
}

function Get-TCPIP6Registry {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline=$true)]
        [string]$Server = $env:COMPUTERNAME
    )

    begin{}

    process {
        $reg = $key = $null
        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
            $key = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\')
            if (-not $key) {
                throw "OpenSubKey failed for 'SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\' on $Server"
            }

            $disabledComponents = $key.GetValue('DisabledComponents','')
            New-Object PSCustomObject -Property @{
                ServerName = $Server;
                DisabledComponents = $disabledComponents
            }
        }
        finally {
            if ($key) { $key.Close() }
            if ($reg) { $reg.Close() }
        }
    }

    end{}
}

function Get-SmbConfig {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Server
    )

    $reg = $key = $null
    try {
        # Could use "Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" but it'd require a remote ps session, which might not be availble for E2010/W2k8R2.
        # Thus using the registry API directly.
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
        $key = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\')
        if (-not $key) {
            throw "OpenSubKey failed for 'SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\' on $Server"
        }

        $smb1 = $key.GetValue('SMB1','')
        New-Object PSCustomObject -Property @{
            ServerName = $Server
            SMB1 = $smb1
        }
    }
    finally {
        if ($key) { $key.Close() }
        if ($reg) { $reg.Close() }
    }
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

function Get-ExSetupVersion {
    [CmdletBinding()]
    param(
        $Server
    )

    $exsetupPath = Join-Path $(Get-ExchangeInstallPath -Server $Server) -ChildPath 'Bin\ExSetup.exe'
    $exsetupPath = ConvertTo-UNCPath $exsetupPath -Server $Server
    (Get-ItemProperty $exsetupPath).VersionInfo
}

<#
  Main
#>

if (-not $FromDateTime) {
    $FromDateTime = [datetime]::MinValue
}

if (-not $ToDateTime) {
    $ToDateTime = [datetime]::MaxValue
}

if ($FromDateTime -ge $ToDateTime) {
    throw "Parameter ToDateTime ($ToDateTime) must be after FromDateTime ($FromDateTime)"
}

$cmd = Get-Command "Get-OrganizationConfig" -ErrorAction:SilentlyContinue
if (-not $cmd) {
    throw "Get-OrganizationConfig is not available. Please run with Exchange Remote PowerShell session"
}
$OrgConfig = Get-OrganizationConfig
$OrgName = $orgConfig.Name
$IsExchangeOnline = $orgConfig.LegacyExchangeDN.StartsWith('/o=ExchangeLabs')

# If the path doesn't exist, create it.
if (-not (Test-Path $Path -ErrorAction Stop)) {
    New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
}
$Path = Resolve-Path $Path

# Create a temporary folder to store data
$tempFolder = New-Item $(Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())) -ItemType directory -ErrorAction Stop

# Prepare for logging
# NOTE: until $logPath is defined, don't call Write-Log
$logFileName = "Log.txt"
$logPath = Join-Path -Path $tempFolder.FullName -ChildPath $logFileName

$startDateTime = Get-Date
Write-Log "Organization Name = $OrgName"
Write-Log "Script Version = $version"
Write-Log "COMPUTERNAME = $env:COMPUTERNAME"
Write-Log "IsExchangeOnline = $IsExchangeOnline"

# Log parameters (raw values are in $PSBoundParameters, but want fixed-up values (e.g. Path)
$sb = New-Object System.Text.StringBuilder
foreach ($paramName in $PSBoundParameters.Keys) {
    $var = Get-Variable $paramName -ErrorAction SilentlyContinue
    if ($var) {
        if ($var.Value -is [DateTime]) {
            $sb.Append("$($var.Name):$($var.Value.ToUniversalTime().ToString('o')); ") | Out-Null
        }
        else {
            $sb.Append("$($var.Name):$($var.Value); ") | Out-Null
        }
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
$directAccessCandidates = New-Object System.Collections.Generic.List[object]
foreach ($server in $Servers) {
    # $Server's value might be something like "e2013*" and matches multiple Servers
    $exServers = @(Get-ExchangeServer $server -ErrorAction SilentlyContinue)

    if (-not $exServers.Count) {
        Write-Log "Get-ExchangeServer did not find any Server matching '$server'"
        continue
    }

    foreach ($exServer in $exServers) {
        # Skip Edge servers unless it's the local server.
        if ($exServer.IsEdgeServer -and $env:COMPUTERNAME -ne $exServer.Name) {
            Write-Log "Dropping $($exServer.Name) from directAccessCandidates since it's an Edge server"
            continue
        }

        # Add if it's not a duplicate
        $inDAS = @($directAccessCandidates | Where-Object {$_.Name -eq $exServer.Name}).Count -gt 0
        if (-not $inDAS) {
            $directAccessCandidates.Add($exServer)
        }
    }
}

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
$OrgConfig | Save-Object -Name "OrganizationConfig"
Run Get-AdminAuditLogConfig
Run Get-AvailabilityAddressSpace
Run Get-AvailabilityConfig
Run Get-OrganizationRelationship
Run "Get-ADServerSettings -WarningAction SilentlyContinue"
Run Get-AuthConfig
Run Get-AuthRedirect
Run Get-AuthServer
Run Get-PartnerApplication
Run Get-DomainController
Run Get-IRMConfiguration
Run Get-OfflineAddressBook
# Run Get-OrganizationalUnit
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
Run Get-DAG
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
# Heath-related command are now commented out since rarely needed.
# Run Get-HealthReport -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity -SkipIfNoServers
# Run Get-ServerHealth -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity -SkipIfNoServers
# Run Test-ServiceHealth -Servers:$directAccessServers -SkipIfNoServers

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
# Run 'Get-ThrottlingPolicyAssociation -ResultSize 1000'

# misc
Write-Progress -Activity $collectionActivity -Status:"Misc" -PercentComplete:85
Run Get-MigrationConfig
Run Get-MigrationEndpoint
Run Get-NetworkConnectionInfo -Servers:$directAccessServers -Identifier:Identity -SkipIfNoServers
# Run Get-ProcessInfo -Servers:$directAccessServers -Identifier:TargetMachine -SkipIfNoServers # skipping, because gwmi Win32_Process is collected (see WMI section)
Run Get-OutlookProtectionRule
Run Get-PolicyTipConfig
Run Get-RbacDiagnosticInfo
Run Get-RoleAssignmentPolicy
# RBAC roles & assignments are skippped for now (can be included in future if necessary)
# Run Get-ManagementRole
# Run Get-ManagementRoleAssignment
# Run Get-ManagementScope

Run Get-SearchDocumentFormat
# Run Get-MailboxAuditBypassAssociation # skipping this because it takes time but rarely needed.
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
    Invoke-FIPS -Servers ($directAccessServers | Where-Object {$_.IsE15OrLater -and $_.IsHubTransportServer})
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
Run 'Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_PowerPlan
Run 'Get-WmiObject -Class Win32_PageFileSetting' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_PageFileSetting
Run 'Get-WmiObject -Class Win32_ComputerSystem' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_ComputerSystem
Run 'Get-WmiObject -Class Win32_OperatingSystem' -Servers $directAccessServers -Identifier ComputerName -SkipIfNoServers -PassThru | Save-Object -Name Win32_OperatingSystem
Run "Get-WmiObject -Class Win32_NetworkAdapterConfiguration" -Servers:$directAccessServers -Identifier:ComputerName -SkipIfNoServers -PassThru |
    Where-Object {$_.IPEnabled} | Save-Object -Name Win32_NetworkAdapterConfiguration
Run "Get-WmiObject -Class Win32_Process" -Servers:$directAccessServers -Identifier:ComputerName -SkipIfNoServers -PassThru | Select-Object ProcessName, Path, CommandLine, ProcessId, ServerName | Save-Object -Name Win32_Process

# Get Exsetup version
Run "Get-ExSetupVersion" -Servers $directAccessServers -SkipIfNoServers

Run Get-SmbConfig -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater}) -SkipIfNoServers

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
        Run "Save-ExchangeEventLog -Path:$eventLogPath -IncludeCrimsonLogs" -Servers $directAccessServers -SkipIfNoServers
    }
    else {
        Run "Save-ExchangeEventLog -Path $eventLogPath" -Servers $directAccessServers -SkipIfNoServers
    }
}

# Collect Perfmon Log
if ($IncludePerformanceLog) {
    Write-Progress -Activity $collectionActivity -Status:"Perfmon Logs" -PercentComplete:90
    Run "Save-ExchangeLogging -Path:$(Join-Path $Path 'Perfmon') -FolderPath 'Diagnostics\DailyPerformanceLogs' -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater}) -SkipIfNoServers
}

# Collect IIS Log
if ($IncludeIISLog) {
    Write-Progress -Activity $collectionActivity -Status:"IIS Logs" -PercentComplete:90
    Run "Save-IISLog -Path:$(Join-Path $Path 'IISLog') -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers -SkipIfNoServers
    Run "Save-HttpErr -Path:$(Join-Path $Path 'HTTPERR') -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers -SkipIfNoServers
}

# Collect Exchange logs (e.g. HttpProxy, Ews, Rpc Client Access, etc.)
# With PowerShellv2, empty array is iterated.
if ($IncludeExchangeLog.Count) {
    foreach ($logType in $IncludeExchangeLog) {
        Write-Progress -Activity $collectionActivity -Status:"$logType Logs" -PercentComplete:90
        Run "Save-ExchangeLogging -Path:`"$(Join-Path $Path $logType)`" -FolderPath '$logType' -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers -SkipIfNoServers
    }
}

# Collect Transport logs (e.g. Connectivity, MessageTracking etc.)
if ($IncludeTransportLog.Count) {
    foreach ($logType in $IncludeTransportLog) {
        Write-Progress -Activity $collectionActivity -Status:"$logType Logs" -PercentComplete:90
        Run "Save-TransportLog -Path:`"$(Join-Path $Path $logType)`" -Type:'$logType' -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers -SkipIfNoServers
    }
}

# Collect Exchange Setup logs (Currently not used. If there's any demand, activate it)
if ($IncludeExchangeSetupLog) {
    Write-Progress -Activity $collectionActivity -Status:"Exchange Setup Logs" -PercentComplete:90
    Run "Save-ExchangeSetupLog -Path:$(Join-Path $Path 'ExchangeSetupLog')" -Servers $directAccessServers -SkipIfNoServers
}

# Collect Fast Search ULS logs
if ($IncludeFastSearchLog) {
    Write-Progress -Activity $collectionActivity -Status:"FastSearch Logs" -PercentComplete:90
    Run "Save-FastSearchLog -Path:$(Join-Path $Path FastSearch) -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers -SkipIfNoServers
}

# Save errors
if ($Script:errs.Count) {
    $errPath = Join-Path $Path -ChildPath "Error"
    if (-not (Test-Path errPath)) {
        New-Item $errPath -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Script.errs | Export-Clixml $(Join-Path $errPath "errs.xml") -Depth 5
}

} # end of try for transcript
finally {
    # release transcript file even when script is stopped in the middle.
    if ($transcriptEnabled) {
        Stop-Transcript
    }

    Remove-Runspace
    Write-Log "Total time is $(((Get-Date) - $startDateTime).TotalSeconds) seconds"
    Close-Log
}

Write-Progress -Activity $collectionActivity -Status:"Packing into a zip file" -PercentComplete:95
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
Write-Host "Done!" -ForegroundColor Green