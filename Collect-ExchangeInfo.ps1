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

.PARAMETER SkipZip
Switch to skip creating a ZIP file. If this is not specified, all the output files will be packed intto a zip file.

.PARAMETER SkipAutoUpdate
Switch to skip auto update. Wihtout this switch, the script first checks to see if there is a newer version available in GitHub repository. If so, it downloads and runs it instead.
This is a best-effort and any failure won't stop the script's execution.

.PARAMETER TrustAllCertificates
Switch to suppress certificate check when accessing a remote web server. This is for the aforementioned auto update.
This script does not access any external web site other than its GitHub repository.

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
    [switch]$SkipZip,
    [switch]$SkipAutoUpdate,
    [switch]$TrustAllCertificates
)

$version = "2021-03-24"
#requires -Version 2.0

<#
  Save object(s) to a text file and optionally export to CliXml.
#>
function Save-Object {
    [CmdletBinding()]
    Param(
        #[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Parameter(ValueFromPipeline=$true)]
        $Object,
        $Name,
        [string]$Path = $Script:Path,
        [bool]$WithCliXml = $true,
        $Depth = 4 # depth for Export-CliXml
    )

    begin {
        # This list accumulate objects to support pipeline. This will be lazily initialized.
        $objectList = $null
        [string]$objectName = $Name
    }

    process {
        # Validate the given objects. If valid, collect them in a list.
        # Collected objects are outputted in the END block

        # When explicitly passed, object is actually a list of objects.
        # When passed from pipeline, object is a single object.
        # To deal with this, use foreach.

        foreach ($o in $Object) {
            if ($null -eq $o) {
                continue
            }

            if (-not $objectName) {
                $objectName = $o.GetType().Name
            }

            if ($null -eq $objectList) {
                $objectList = New-Object System.Collections.Generic.List[object]
            }

            $objectList.Add($o)

        }
    }

    end {
        if ($objectList.Count -gt 0) {
            if ($WithCliXml) {
                try {
                    # Export-Clixml could fail for non-CLS-compliant objects
                    $objectList | Export-Clixml -Path:([IO.Path]::Combine($Path, "$objectName.xml")) -Encoding:UTF8 -Depth $Depth
                }
                catch {
                    Write-Error "Export-CliXml failed. $_"
                }
            }

            $objectList | Format-List * | Out-File ([IO.Path]::Combine($Path, "$objectName.txt")) -Encoding:UTF8
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
        Write-Verbose "System.IO.Compression.FileSystem wasn't found. Using alternate method"
    }

    $files = @(Get-ChildItem $Path -Recurse | Where-Object {-not $_.PSIsContainer})

    # Apply filters
    if ($FromDateTime)  {
        $files = @($files | Where-Object {$_.LastWriteTime -ge $FromDateTime})
    }

    if ($ToDateTime) {
        $files = @($files | Where-Object {$_.LastWriteTime -le $ToDateTime})
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
            $prevProgress = 0

            foreach ($file in $files) {
                $progress = 100 * $count / $files.Count
                if ($progress -gt $prevProgress + 10) {
                    Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Please wait" -PercentComplete $progress
                    $prevProgress = $progress
                }

                $fileStream = $zipEntryStream = $null
                try {
                    $fileStream = New-Object System.IO.FileStream -ArgumentList $file.FullName, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::ReadWrite)
                    $zipEntry = $zipArchive.CreateEntry($file.FullName.Substring($Path.Length + 1))
                    $zipEntryStream = $zipEntry.Open()
                    $fileStream.CopyTo($zipEntryStream)

                    ++$count
                }
                catch {
                    Write-Error -Message "Failed to add $($file.FullName). $_" -Exception $_.Exception
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
        if ($null -eq $FromDateTime -and $null -eq $ToDateTime) {
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
            Write-Error -Message "Cannot create a zip file on $Server. Each log file will be copied. $_" -Exception $_.Exception
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
        $files = @(Get-ChildItem $SourcePath -Recurse -Filter $Filter | Where-Object {-not $_.PSIsContainer})

        if ($FromDateTime)  {
            $files = @($files | Where-Object {$_.LastWriteTime -ge $FromDateTime})
        }

        if ($ToDateTime) {
            $files = @($files | Where-Object {$_.LastWriteTime -le $ToDateTime})
        }

        foreach ($file in $files) {
            $destination = Join-Path $DestitionPath $file.DirectoryName.SubString($SourcePath.Length)
            if (-not (Test-Path $destination)) {
                New-Item $destination -ItemType Directory | Out-Null
            }

            try {
                Copy-Item $file.FullName -Destination $destination -Force
            }
            catch {
                Write-Error $_
            }
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
                # Flush the log buffer
                netsh http flush logbuffer | Out-Null

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
        # Give some time to flush log data.
        Start-Sleep -Seconds 5

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
    $logPath = [IO.Path]::Combine($env:SystemRoot, 'System32\LogFiles\HTTPERR')

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

    # Default path: %ExchangeInstallPath% + $FolderPath
    $logPath = $null
    $exchangePath = Get-ExchangeInstallPath -Server $Server
    if ($exchangePath) {
        $logPath = [IO.Path]::Combine($exchangePath, "Logging\$FolderPath")
    }

    # Diagnostics path can be modified. So update the folder path if necessary
    if ($FolderPath -like 'Diagnostics\*') {
        $customPath = $null
        try {
            $customPath = Get-DiagnosticsPath -Server $Server -ErrorAction SilentlyContinue
        }
        catch {
            Write-Error -Message "Get-DiagnosticsPath failed. $_." -Exception $_.Exception
        }

        if ($customPath) {
            $subPath = $FolderPath.Substring($FolderPath.IndexOf('\') + 1)
            $logPath = [IO.Path]::Combine($customPath, $subPath)
            Write-Log "[$($MyInvocation.MyCommand)] Custom Diagnostics path is found. Using $logPath"
        }
    }

    if (-not $logPath) {
        Write-Error "Cannot fine the target log path."
        return
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
        $transport = Get-TransportService $Server -ErrorAction SilentlyContinue
    }
    elseif (Get-Command 'Get-TransportServer' -ErrorAction SilentlyContinue) {
        $transport = Get-TransportServer $Server -ErrorAction SilentlyContinue
    }

    # If both Get-TransportService & Get-TransportServer are not available, bail.
    if (-not $transport) {
        throw "Get-TransportService/TransportServer is not available."
    }

    $frontendTransport = $null
    if (Get-Command 'Get-FrontendTransportService' -ErrorAction SilentlyContinue) {
        $frontendTransport = Get-FrontendTransportService $Server -ErrorAction SilentlyContinue
    }

    # Before saving, try to flush the logs. This is a best effort.
    # Sending control code 206 should flush the logs.
    $flushSuccess = $false
    foreach ($service in @('MSExchangeTransport', 'MSExchangeFrontEndTransport')) {
        $serviceController = Get-Service $service -ComputerName $Server -ErrorAction SilentlyContinue
        if ($serviceController) {
            $err = $($serviceController.ExecuteCommand(206)) 2>&1
            $serviceController.Dispose()
            if (-not $flushSuccess -and $null -eq $err) {
                $flushSuccess = $true
            }
        }
    }

    # Flush request was successful for at least one of services. So wait a little to give time to flush log data.
    if ($flushSuccess) {
        Start-Sleep -Seconds 5
    }

    foreach ($logType in $Type) {
        # Parameter name is ***LogPath
        $paramName = $logType + 'LogPath'
        if (-not $transport.$paramName) {
            Write-Error "Cannot find $paramName in the result of Get-TransportService"
            continue
        }
        $sourcePath = ConvertTo-UNCPath $transport.$paramName.ToString() -Server $Server
        $destination = Join-Path $Path "$logType\$Server\Hub"
        Save-Item -SourcePath $sourcePath -DestitionPath $destination -FromDateTime $FromDateTime -ToDateTime $ToDateTime

        if ($frontendTransport -and $frontendTransport.$paramName) {
            $sourcePath = ConvertTo-UNCPath $frontendTransport.$paramName.ToString() -Server $Server
            $destination = Join-Path $Path "$logType\$Server\FrontEnd"
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

    $exsetupPath = Get-ExchangeInstallPath -Server $Server -ErrorAction Stop
    $source = ConvertTo-UNCPath $([IO.Path]::Combine($exsetupPath, 'Bin\Search\Ceres\Diagnostics\Logs')) -Server $Server
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
    if (-not (Get-Command 'ldifde.exe' -ErrorAction SilentlyContinue -ErrorVariable err)) {
        Write-Error "Ldifde is not available. $err"
        return
    }

    if ($Script:OrgConfig) {
        $exorg = $Script:OrgConfig.DistinguishedName
    }
    else {
        $exorg = (Get-OrganizationConfig).DistinguishedName
    }

    if (-not $exorg) {
        Write-Error "Couldn't get Exchange org DN"
        return
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
        try {
            $_.GetType() | Out-Null
            $_.Exception.GetType() | Out-Null
            Write-Log "[Terminating Error] '$Command' failed. $($_.ToString()) $(if ($_.Exception.Line) {"(At line:$($_.Exception.Line) char:$($_.Exception.Offset))"})"
            if ($null -ne $Script:errs) {$Script.errs.Add($_)}
        }
        catch {
            Write-Log "$Command threw a non-CLS-compliant exception object."
        }
    }
    finally {
        if ($errs.Count) {
            foreach ($err in $errs) {
                try {
                    $err.GetType() | Out-Null
                    $err.Exception.GetType() | Out-Null
                    Write-Log "[Non-Terminating Error] Error in '$Command'. $($err.ToString()) $(if ($err.Exception.Line) {"(At line:$($err.Exception.Line) char:$($err.Exception.Offset))"})"
                }
                catch {
                    Write-Log "$Command returned a non-CLS-compliant error object. $err"
                }
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
        [Parameter(ValueFromPipeline=$true)]
        [object[]]$ResultCollection,
        [switch]$RemoveDuplicate,
        [switch]$PassThru,
        [int]$TimeoutSeconds = 180
    )

    begin {
        # This will hold both the pipelined objects and output objects. This will be lazily initialized.
        $result = $null
    }

    # Accumulate the previous results
    process {
        # Make sure not to add $null and collection itself
        foreach ($pipedObj in $ResultCollection) {
            # In PowerShellV2, $null is iterated over.
            if ($pipedObj) {
                if ($null -eq $result) {
                    $result = New-Object System.Collections.Generic.List[object]
                }
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
            if (-not $PSBoundParameters.ContainsKey('Servers')) {
                RunCommand $Command -TimeoutSeconds $TimeoutSeconds
            }
            elseif ($Servers.Count) {
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

        # On a rare situation, you might get non-CLS-compliant objects. Trying to access a property causes a terminating error. So filter them out.
        # For example, Get-MailboxDatabaseCopyStatus could return "Microsoft.Exchange.Cluster.Replay.FailedToOpenLogTruncContextException", which is not CLS-compliant.
        $temp = @(
            for ($i = 0; $i -lt $temp.Count; ++$i) {
                try {
                    # If GetType() fails, most likely this type is not CLS-compliant
                    $temp[$i].GetType() | Out-Null
                }
                catch {
                    Write-Log "$Command returned a non-CLS-compliant type"
                    continue
                }

                $temp[$i]
            }
        )

        # Deserialize if SerializationData property is available.
        if (-not $Script:formatter) {
            $Script:formatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        }

        for ($i = 0; $i -lt $temp.Count; ++$i) {
            if ($null -ne $temp[$i].serializationData) {
                try {
                    $stream = New-Object System.IO.MemoryStream -ArgumentList (, $temp[$i].serializationData)
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

        # Shortcut. If there was no pipelined objects and no output at this point, there's nothing else to do.
        if (-not $result.Count -and -not $temp.Count) {
            return
        }

        if (-not $RemoveDuplicate) {
            if ($null -eq $result) {
                $result = $temp
            }
            else {
                $result.AddRange($temp)
            }
        }
        else {
            if ($null -eq $result) {
                $result = New-Object System.Collections.Generic.List[object]
            }

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
                Run "$($command.Name) -ShowMailboxVirtualDirectories" -Servers:($allExchangeServers | Where-Object {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -RemoveDuplicate -PassThru |
                    Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -RemoveDuplicate
            }
            else {
                Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -RemoveDuplicate
            }
        }
        else {
            if ($IncludeIISVirtualDirectories) {
                Run "$($command.Name)" -Servers:($allExchangeServers | Where-Object {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -RemoveDuplicate -PassThru |
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
                    foreach ($err in $errs) {
                        Write-Log "[Non-Terminiating Error]$err"
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

    # Check if setspn.exe is available
    if (-not (Get-Command 'setspn.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "setspn.exe is not available"
        return
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
    $winTempEventPath = [IO.Path]::Combine($winTempPath, "EventLogs_$(Get-Date -Format "yyyyMMdd_HHmmss")")
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
        $localFilePath = [IO.Path]::Combine($winTempEventPath, $fileName)
        wevtutil epl $log $localFilePath /ow /r:$Server
        wevtutil al $localFilePath /r:$Server
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

    [IO.Path]::Combine($win32os.WindowsDirectory, 'Temp')
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

    $exchangePath = $win32env | Where-Object {$_.Name -eq 'ExchangeInstallPath'} | Select-Object -First 1
    if (-not $exchangePath.VariableValue) {
        Write-Error "Cannot find ExchangeInstallPath on $Server"
        return
    }

    $exchangePath.VariableValue
}

function Get-DAG {
    [CmdletBinding()]
    param()

    $dags = @(RunCommand Get-DatabaseAvailabilityGroup)

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
                    if ($null -eq $versionKeyName) {continue}

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
                            if ($null -eq $subKeyName) {continue}

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
                        if ($null -eq $subKeyName) { continue }

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

function Get-FipsAlgorithmPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Server
    )

    begin{}

    process{
        $hklm = $key = $null
        try {
            $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Server)
            $key = $hklm.OpenSubKey('SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy\')

            if (-not $key) {
                Write-Error "OpenSubKey failed for'SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy\' on $Server"
                return
            }

            $enabled = $key.GetValue('Enabled', 0)

            New-Object PSCustomObject -Property @{
                ServerName = $Server
                Enabled = $enabled -ne 0
            }

        }
        catch {
            Write-Error -ErrorRecord $_
        }
        finally {
            if ($key) { $key.Close() }
            if ($hklm) {$hklm.Close() }
        }
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
            Write-Error -Message "Failed to invoke command on a remote session to $Server.$_" -Exception $_.Exception
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

    $exsetupPath = [IO.Path]::Combine($(Get-ExchangeInstallPath -Server $Server -ErrorAction Stop), 'Bin\ExSetup.exe')
    $exsetupPath = ConvertTo-UNCPath $exsetupPath -Server $Server
    (Get-ItemProperty $exsetupPath).VersionInfo
}

function Get-ProxySettingInternal {
    [CmdletBinding()]
    param(
    )

    $props = @{}

    # Use Win32 WinHttpGetDefaultProxyConfiguration
    # I'm not using "netsh winhttp show proxy", because the output is system language dependent.  Netsh just calls this function anyway.
    $WinHttpDef = @'
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct WINHTTP_PROXY_INFO
{
    public uint dwAccessType;
    public string lpszProxy;
    public string lpszProxyBypass;
}

// From winhttp.h
// WinHttpOpen dwAccessType values (also for WINHTTP_PROXY_INFO::dwAccessType)
public enum ProxyAccessType
{
    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0,
    WINHTTP_ACCESS_TYPE_NO_PROXY = 1,
    WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3,
    WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY = 4
}

[DllImport("winhttp.dll", SetLastError = true)]
public static extern bool WinHttpGetDefaultProxyConfiguration(out WINHTTP_PROXY_INFO proxyInfo);
'@

    if (-not ('Win32.WinHttp' -as [type])) {
        Add-Type -MemberDefinition $WinHttpDef -Name WinHttp -Namespace Win32
    }

    $proxyInfo = New-Object Win32.WinHttp+WINHTTP_PROXY_INFO
    if ([Win32.WinHttp]::WinHttpGetDefaultProxyConfiguration([ref] $proxyInfo)) {
        $props['WinHttpDirectAccess'] = $proxyInfo.dwAccessType -eq [Win32.WinHttp+ProxyAccessType]::WINHTTP_ACCESS_TYPE_NO_PROXY
        $props['WinHttpProxyServer'] = $proxyInfo.lpszProxy
        $props['WinHttpBypassList'] = $proxyInfo.lpszProxyBypass
        $props['WINHTTP_PROXY_INFO'] = $proxyInfo # for debugging purpuse
    }
    else {
        Write-Error ("Win32 WinHttpGetDefaultProxyConfiguration failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    Write-Verbose "WinHttp*** properties correspond to WINHTTP_PROXY_INFO obtained by WinHttpGetDefaultProxyConfiguration. See https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_current_user_ie_proxy_config"
    New-Object PSCustomObject -Property $props
}


function Get-ProxySetting {
    [CmdletBinding()]
    param(
    [Alias('ComputerName')]
    [string]$Server = $env:COMPUTERNAME
    )

    if ($env:COMPUTERNAME -eq $Server) {
        Get-ProxySettingInternal
        return
    }

    $session = $null
    try {
        $session = New-PSSession -ComputerName $Server -ErrorAction SilentlyContinue
        if (-not $session) {
            Write-Error "Cannot make a PSSession to $Server."
            return
        }

        Invoke-Command -Session $session -ScriptBlock ${Function:Get-ProxySettingInternal}
    }
    finally {
        if ($session) {
            Remove-PSSession $session
        }
    }
}

function Get-NetworkInterface {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory =  $true)]
        $Server
    )

    $scriptBlock = {
        $nics = [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()
        foreach ($nic in $nics) {
            # Exract just the properties
            $adapter = @{}
            foreach ($prop in @($nic | Get-Member -MemberType Properties)) {
                $adapter[$prop.Name] = $nic.$($prop.Name)
            }

            # Create another object from GetIPProperties() and embed it in the adapter object
            # I need this because otherwise properties like UnicastAddresses become a plain string object.
            $ipInfo = $nic.GetIPProperties()
            $IPProperties = @{}
            foreach ($prop in @($ipInfo | Get-Member -MemberType Properties)) {
                $IPProperties[$prop.Name] = $ipInfo.$($prop.Name)
            }

            $adapter['IPProperties'] = New-Object PSCustomObject -Property $IPProperties

            # This is the final object to return
            New-Object PSCustomObject -property $adapter
        }
    }

    if ($env:COMPUTERNAME -eq $Server) {
        Invoke-Command -ScriptBlock $scriptBlock
    }
    else {
        $session = $null
        try {
            $session = New-PSSession -ComputerName $Server
            if ($session) {
                Invoke-Command -Session $session -ScriptBlock $scriptBlock
            }
        }
        finally {
            if ($session) {
                Remove-PSSession $session
            }
        }
    }
}

<#
Check the state of Transport's UnifiedContent folder.
#>
function Get-UnifiedContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Server
    )

    # Only E2013 or later.
    $exServer = Get-ExchangeServer $Server
    if (-not $exServer.IsE15OrLater) {
        Write-Log "Skipping $Server because it is E2010 or before."
        return
    }

    # Find Transport's TemporaryStoragePath from config file.
    $edgeConfigFile = [IO.Path]::Combine($(Get-ExchangeInstallPath -Server $Server -ErrorAction Stop), 'bin\EdgeTransport.exe.config')
    $edgeConfigFileUNC = ConvertTo-UNCPath -Server $Server.ToString() -Path $edgeConfigFile

    $reader = $null
    try {
        $reader = [IO.File]::OpenText($edgeConfigFileUNC)
        $tempPath = $null
        while ($line = $reader.ReadLine()) {
            if ($line -match '<add key="TemporaryStoragePath" +value="(?<tempPath>.+)"') {
                $tempPath = $Matches['tempPath']
                break
            }
        }

        if (-not $tempPath) {
            Write-Error "Cannot find TemporaryStoragePath in $edgeConfigFileUNC"
            return
        }

        $tempPathUNC = ConvertTo-UNCPath -Server $Server.ToString() -Path $tempPath
        $unifiedContent = Join-Path $tempPathUNC 'UnifiedContent'

        $totalSize = 0
        $count = 0

        Get-ChildItem $unifiedContent | ForEach-Object {
            $totalSize += $_.Length
            $count ++
        }

        New-Object PSCustomObject -Property @{
            Server = $Server.ToString()
            TemporaryStoragePath = $tempPath
            TotalBytes = $totalSize
            Count = $count
        }
    }
    finally {
        if ($reader) {
            $reader.Dispose()
        }
    }
}

<#
Save Exchange's application config files
#>
function Save-AppConfig {
    [CmdletBinding()]
    param(
        # Where to save the file
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $Server
    )

    $exServer = Get-ExchangeServer $Server
    if (-not $exServer.IsHubTransportServer) {
        Write-Log "Skipping $Server because this is not a HubTransportServer."
        return
    }

    $Folder = Join-Path $Path $Server
    if (-not (Test-Path $Folder)) {
        New-Item -ItemType Directory $Folder | Out-Null
    }

    $exchangePath = Get-ExchangeInstallPath -Server $Server -ErrorAction Stop
    $binFolder = [IO.Path]::Combine($exchangePath, 'bin')
    $binFolderUNC = ConvertTo-UNCPath -Server $Server -Path $binFolder

    Save-Item -SourcePath $binFolderUNC -DestitionPath $Folder -Filter '*.exe.config' -SkipZip

    # For now, web config files are not included.
    # $casFolder = [IO.Path]::Combine($exchangePath, 'ClientAccess')
    # $casFolderUNC = ConvertTo-UNCPath -Server $Server -Path $casFolder
    # Save-Item -SourcePath $casFolderUNC -DestitionPath (Join-Path $Folder 'ClientAccess') -Filter 'web.config' -SkipZip
}

function Get-InstalledUpdate
{
    [CmdletBinding()]
    param(
    [string]$Server = $env:COMPUTERNAME
    )

    function Get-InstalledUpdateInternal
    {
        [CmdletBinding()]
        param()

        # Ask items in AppUpdatesFolder from Shell
        # FOLDERID_AppUpdates == a305ce99-f527-492b-8b1a-7e76fa98d6e4
        $shell = $appUpdates = $null

        try {
            $shell = New-Object -ComObject Shell.Application
            $appUpdates = $shell.NameSpace('Shell:AppUpdatesFolder')
            if ($null -eq $appUpdates) {
                Write-Log "Cannot obtain Shell:AppUpdatesFolder. Probabliy 32bit PowerShell is used on 64bit OS"
                Write-Error "Cannot obtain Shell:AppUpdatesFolder"
                return
            }

            $items = $appUpdates.Items()

            foreach ($item in $items) {
                # https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
                New-Object PSCustomObject -Property @{
                    Name        = $item.Name
                    Program     = $appUpdates.GetDetailsOf($item, 2)
                    Version     = $appUpdates.GetDetailsOf($item, 3)
                    Publisher   = $appUpdates.GetDetailsOf($item, 4)
                    URL         = $appUpdates.GetDetailsOf($item, 7)
                    InstalledOn = $appUpdates.GetDetailsOf($item, 12)
                }
                [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($item) | Out-Null
            }
        }
        finally {
            if ($appUpdates) {
                [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($appUpdates) | Out-Null
            }
            if ($shell) {
                [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shell) | Out-Null
            }
        }
    }

    if ($Server -eq $env:COMPUTERNAME) {
        Get-InstalledUpdateInternal
        return
    }

    $session = $null

    try {
        $session = New-PSSession -ComputerName $Server -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock ${Function:Get-InstalledUpdateInternal}
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        if ($session) {
            Remove-PSSession $session
        }
    }
}

function Get-NLMConnectivity {
    [CmdletBinding()]
    param()

    $CLSID_NetworkListManager = [Guid]'DCB00C01-570F-4A9B-8D69-199FDBA5723B'
    $type = [Type]::GetTypeFromCLSID($CLSID_NetworkListManager)
    $nlm = [Activator]::CreateInstance($type)

    $isConnectedToInternet = $nlm.IsConnectedToInternet
    $conn = $nlm.GetConnectivity()

    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nlm) | Out-Null
    $nlm = $null

    # NLM_CONNECTIVITY enumeration
    # https://docs.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_connectivity

    # From netlistmgr.h
    $NLM_CONNECTIVITY = @{
        NLM_CONNECTIVITY_DISCONNECTED      = 0
        NLM_CONNECTIVITY_IPV4_NOTRAFFIC    = 1
        NLM_CONNECTIVITY_IPV6_NOTRAFFIC    = 2
        NLM_CONNECTIVITY_IPV4_SUBNET	   = 0x10
        NLM_CONNECTIVITY_IPV4_LOCALNETWORK = 0x20
        NLM_CONNECTIVITY_IPV4_INTERNET	   = 0x40
        NLM_CONNECTIVITY_IPV6_SUBNET	   = 0x100
        NLM_CONNECTIVITY_IPV6_LOCALNETWORK = 0x200
        NLM_CONNECTIVITY_IPV6_INTERNET	   = 0x400
    }

    $connectivity = New-Object System.Collections.Generic.List[string]

    foreach ($entry in $NLM_CONNECTIVITY.GetEnumerator()) {
        if ($conn -band $entry.Value) {
            $connectivity.Add($entry.Key)
        }
    }

    [PSCustomObject]@{
        IsConnectedToInternet = $isConnectedToInternet
        Connectivity = $connectivity
    }
}

<#
Check GitHub's latest release and if it's newer, download and import it except if OutlookTrace is installed as module.
#>
function Invoke-AutoUpdate {
    [CmdletBinding()]
    param(
        [uri]$GitHubUri = 'https://api.github.com/repos/jpmessaging/CollectExchangeInfo/releases/latest'
    )

    $autoUpdateSuccess = $false
    $message = $null

    if (-not (Get-Command 'Invoke-WebRequest' -ErrorAction SilentlyContinue)) {
        $message = "Skipped autoupdate because Invoke-WebRequest is not available (Probably running with PSv2)."
    }
    elseif (-not (Get-NLMConnectivity).IsConnectedToInternet) {
        $message = "Skipped autoupdate because there's no connectivity to internet."
    }
    else {
        try {
            Write-Progress -Activity "AutoUpdate" -Status 'Checking if a newer version is available. Please wait' -PercentComplete -1
            $release = Invoke-RestMethod -Uri $GitHubUri -ErrorAction Stop

            # release.name may look like "v2020-10-09". Extrace just the date.
            $latestVersion = $release.name
            if ($release.name -match '\d{4}-\d{2}-\d{2}') {
                $latestVersion = $Matches[0]
            }

            if ($Version -ge $latestVersion) {
                $message = "Skipped because the current script ($Version) is newer than GitHub's latest release ($($release.name))."
            }
            else {
                Write-Verbose "Downloading the latest script."
                $response = Invoke-Command {
                    # Suppress progress on Invoke-WebRequest.
                    $ProgressPreference = "SilentlyContinue"
                    Invoke-WebRequest -Uri $release.assets.browser_download_url -UseBasicParsing
                }

                # Rename the current script and replace with the latest one.
                $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + "_" + $Version + [IO.Path]::GetExtension($PSCommandPath)
                if (Test-Path (Join-Path ([IO.Path]::GetDirectoryName($PSCommandPath)) $newName)) {
                    $newName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath) + "_" + $Version + [IO.Path]::GetRandomFileName() + [IO.Path]::GetExtension($PSCommandPath)
                }

                Rename-Item -LiteralPath $PSCommandPath -NewName $newName -ErrorAction Stop
                [IO.File]::WriteAllBytes($PSCommandPath, $response.Content)

                Write-Verbose "Lastest script ($($release.name)) was successfully downloaded."
                $autoUpdateSuccess = $true
            }
        }
        catch {
            $message = "Autoupdate failed. $_"
        }
        finally {
            Write-Progress -Activity "AutoUpdate" -Status "done" -Completed
        }
    }
        New-Object PSCustomObject -Property @{
            Success = $autoUpdateSuccess
            Message = $message
        }
}


<#
  Main
#>

# This is just for testing.
$TrustAllCertificatePolicyDefinition = @"
using System.Net;
using System.Security.Cryptography.X509Certificates;

public class TrustAllCertsPolicy : ICertificatePolicy
{
    public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem)
    {
        return true;
    }
}
"@

if ($TrustAllCertificates) {
    if (-not ("TrustAllCertsPolicy" -as [type])) {
        Add-Type $TrustAllCertificatePolicyDefinition
    }

    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

# Check if a new version is available and use it if possible. This is just a best effort thing.
if (-not $SkipAutoUpdate) {
    $autoUpdate = Invoke-AutoUpdate
    if ($autoUpdate.Success) {
        $updatedSelf = Get-Command $PSCommandPath

        # Get the list of current parameters that's also available in the updated cmdlet
        $params = @{}
        foreach ($currentParam in $PSBoundParameters.GetEnumerator()) {
            if ($updatedSelf.Parameters.ContainsKey($currentParam.Key)) {
                $params.Add($currentParam.Key, $currentParam.Value)
            }
        }

        if ($updatedSelf.Parameters.ContainsKey('SkipAutoUpdate')) {
            $params.Add('SkipAutoUpdate', $true)
        }

        & $updatedSelf @params
        return
    }
}

if (-not $FromDateTime) {
    $FromDateTime = [DateTime]::MinValue
}

if (-not $ToDateTime) {
    $ToDateTime = [DateTime]::MaxValue
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

Write-Log "AutoUpdate: $(if ($SkipAutoUpdate) {'Skipped due to SkipAutoUpdate switch'} else {$autoUpdate.Message})"

# Log parameters (raw values are in $PSBoundParameters, but want fixed-up values (e.g. Path)
$sb = New-Object System.Text.StringBuilder
foreach ($paramName in $PSBoundParameters.Keys) {
    $var = Get-Variable $paramName -ErrorAction SilentlyContinue
    if ($var) {
        if ($var.Value -is [DateTime]) {
            $sb.Append("$($var.Name):$($var.Value.ToUniversalTime().ToString('o')); ") | Out-Null
        }
        else {
            $sb.Append("$($var.Name):$($var.Value -join ','); ") | Out-Null
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
    Start-Transcript -Path $transcriptPath -NoClobber -ErrorAction:Stop | Out-Null
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
Run Get-AuthenticationPolicy
Run Get-ClientAccessRule
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
Run Get-IPAllowListEntry -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsHubTransportServer})
Run Get-IPAllowListProvider
Run Get-IPAllowListProvidersConfig
Run Get-IPBlockListConfig
Run Get-IPBlockListEntry -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsHubTransportServer})
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

# Server Settings
Write-Progress -Activity $collectionActivity -Status:"Server Settings" -PercentComplete:40
Run Get-ExchangeServer
Run Get-MailboxServer

# For CAS (>= E14) in DAS list, include ASA info
Run "Get-ClientAccessServer -IncludeAlternateServiceAccountCredentialStatus -WarningAction:SilentlyContinue" -Servers:($allExchangeServers | Where-Object {$_.IsDirectAccess -and $_.IsClientAccessServer -and -$_.IsE14OrLater}) -Identifier:Identity -RemoveDuplicate -PassThru |
    Run "Get-ClientAccessServer -WarningAction:SilentlyContinue" -Identifier:Identity -RemoveDuplicate

Run Get-ClientAccessArray
Run Get-RpcClientAccess
Run "Get-TransportServer -WarningAction:SilentlyContinue"
Run Get-TransportService
Run Get-FrontendTransportService
Run Get-ExchangeDiagnosticInfo -Servers $directAccessServers
Run Get-ExchangeServerAccessLicense

Run Get-PopSettings -Servers:$allExchangeServers
Run Get-ImapSettings -Servers:$allExchangeServers

Write-Log "Server Done"

# Database
Write-Progress -Activity $collectionActivity -Status:"Database Settings" -PercentComplete:50

Run "Get-MailboxDatabase -Status -IncludePreExchange" -Servers:($allExchangeServers | Where-Object {$_.IsMailboxServer -and $_.IsDirectAccess}) -RemoveDuplicate -PassThru |
    Run "Get-MailboxDatabase -IncludePreExchange" -RemoveDuplicate

Run "Get-PublicFolderDatabase -Status" -Servers:($allExchangeServers | Where-Object {$_.IsMailboxServer -and $_.IsDirectAccess}) -RemoveDuplicate -PassThru |
    Run "Get-PublicFolderDatabase" -RemoveDuplicate

Run Get-MailboxDatabaseCopyStatus -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater -and $_.IsMailboxServer})
Run Get-DAG
Run Get-DatabaseAvailabilityGroupConfiguration
if (Get-Command Get-DatabaseAvailabilityGroup -ErrorAction:SilentlyContinue) {
    Run "Get-DatabaseAvailabilityGroupNetwork -ErrorAction:SilentlyContinue" -Servers:(Get-DatabaseAvailabilityGroup) -Identifier:'Identity'
}
Write-Log "Database Done"

# Virtual Directories
Write-Progress -Activity $collectionActivity -Status:"Virtual Directory Settings" -PercentComplete:60
Run 'Get-VirtualDirectory'
Run "Get-IISWebBinding" -Servers $directAccessServers -PassThru | Save-Object -Name WebBinding

# Active Monitoring & Managed Availability
Write-Progress -Activity $collectionActivity -Status:"Monitoring Settings" -PercentComplete:70
Run Get-GlobalMonitoringOverride
Run Get-ServerMonitoringOverride -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater})
Run Get-ServerComponentState -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity
# Heath-related command are now commented out since rarely needed.
# Run Get-HealthReport -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity
# Run Get-ServerHealth -Servers:($directAccessServers | Where-Object {$_.IsE15OrLater}) -Identifier:Identity
# Run Test-ServiceHealth -Servers:$directAccessServers

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
Run Get-ExchangeCertificate -Servers:($directAccessServers | Where-Object {$_.IsE14OrLater})

# Throttling
Write-Progress -Activity $collectionActivity -Status:"Throttling" -PercentComplete:85
Run Get-ThrottlingPolicy
# Run 'Get-ThrottlingPolicyAssociation -ResultSize 1000'

# misc
Write-Progress -Activity $collectionActivity -Status:"Misc" -PercentComplete:85
Run Get-MigrationConfig
Run Get-MigrationEndpoint
Run Get-NetworkConnectionInfo -Servers:$directAccessServers -Identifier:Identity
# Run Get-ProcessInfo -Servers:$directAccessServers -Identifier:TargetMachine # skipping, because gwmi Win32_Process is collected (see WMI section)
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
Run Get-DotNetVersion -Servers:($directAccessServers) -Identifier:Server

# TLS Settings
Run Get-TlsRegistry -Servers $directAccessServers -Identifier:Server

# TCPIP6
Run Get-TCPIP6Registry -Servers $directAccessServers -Identifier:Server

# MSInfo32
# Get-MSInfo32 -Servers $directAccessServers

Run Get-ProxySetting -Servers $directAccessServers
Run Get-NetworkInterface -Server $directAccessServers

# WMI
# Win32_powerplan is available in Win7 & above.
Run 'Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan' -Servers $directAccessServers -Identifier ComputerName -PassThru | Save-Object -Name Win32_PowerPlan
Run 'Get-WmiObject -Class Win32_PageFileSetting' -Servers $directAccessServers -Identifier ComputerName -PassThru | Save-Object -Name Win32_PageFileSetting
Run 'Get-WmiObject -Class Win32_ComputerSystem' -Servers $directAccessServers -Identifier ComputerName -PassThru | Save-Object -Name Win32_ComputerSystem
Run 'Get-WmiObject -Class Win32_OperatingSystem' -Servers $directAccessServers -Identifier ComputerName -PassThru | Save-Object -Name Win32_OperatingSystem
Run "Get-WmiObject -Class Win32_NetworkAdapterConfiguration" -Servers:$directAccessServers -Identifier:ComputerName -PassThru |
    Where-Object {$_.IPEnabled} | Save-Object -Name Win32_NetworkAdapterConfiguration
Run "Get-WmiObject -Class Win32_Process" -Servers:$directAccessServers -Identifier:ComputerName -PassThru | Select-Object ProcessName, Path, CommandLine, ProcessId, ServerName | Save-Object -Name Win32_Process

# Get Exsetup version
Run "Get-ExSetupVersion" -Servers $directAccessServers

Run Get-SmbConfig -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater})
Run Get-FipsAlgorithmPolicy -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater})
Run "Save-AppConfig -Path $(Join-Path $Path 'AppConfig')" -Servers $directAccessServers
Run Get-UnifiedContent -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater})
# Run Get-InstalledUpdate -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater})

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

    $eventLogPath = Join-Path $Path -ChildPath 'EventLog'
    if ($IncludeEventLogsWithCrimson) {
        Run "Save-ExchangeEventLog -Path:$eventLogPath -IncludeCrimsonLogs" -Servers $directAccessServers
    }
    else {
        Run "Save-ExchangeEventLog -Path $eventLogPath" -Servers $directAccessServers
    }
}

# Collect Perfmon Log
if ($IncludePerformanceLog) {
    Write-Progress -Activity $collectionActivity -Status:"Perfmon Logs" -PercentComplete:90
    Run "Save-ExchangeLogging -Path:$(Join-Path $Path 'Perfmon') -FolderPath 'Diagnostics\DailyPerformanceLogs' -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $($directAccessServers | Where-Object {$_.IsE15OrLater})
}

# Collect IIS Log
if ($IncludeIISLog) {
    Write-Progress -Activity $collectionActivity -Status:"IIS Log" -PercentComplete:90
    Run "Save-IISLog -Path:$(Join-Path $Path 'IISLog') -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers
    Run "Save-HttpErr -Path:$(Join-Path $Path 'HTTPERR') -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers
}

# Collect Exchange logs (e.g. HttpProxy, Ews, Rpc Client Access, etc.)
foreach ($logType in $IncludeExchangeLog) {
    if (-not $logType) { continue } # With PowerShellv2, $null is iterated.
    Write-Progress -Activity $collectionActivity -Status:"$logType Logs" -PercentComplete:90
    Run "Save-ExchangeLogging -Path:`"$(Join-Path $Path $logType)`" -FolderPath '$logType' -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers
}

# Collect Transport logs (e.g. Connectivity, MessageTracking etc.)
if ($IncludeTransportLog.Count) {
    Write-Progress -Activity $collectionActivity -Status:"Transport Logs" -PercentComplete:90
    Run "Save-TransportLog -Path:`"$(Join-Path $Path 'TransportLog')`" -Type:$($IncludeTransportLog -join ',') -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers
}

# Collect Exchange Setup logs (Currently not used. If there's any demand, activate it)
if ($IncludeExchangeSetupLog) {
    Write-Progress -Activity $collectionActivity -Status:"Exchange Setup Logs" -PercentComplete:90
    Run "Save-ExchangeSetupLog -Path:$(Join-Path $Path 'ExchangeSetupLog')" -Servers $directAccessServers
}

# Collect Fast Search ULS logs
if ($IncludeFastSearchLog) {
    Write-Progress -Activity $collectionActivity -Status:"FastSearch Logs" -PercentComplete:90
    Run "Save-FastSearchLog -Path:$(Join-Path $Path FastSearchLog) -FromDateTime:'$FromDateTime' -ToDateTime:'$ToDateTime'" -Servers $directAccessServers
}

# Save errors
if ($Script:errs.Count) {
    $errPath = Join-Path $Path -ChildPath "Error"
    if (-not (Test-Path errPath)) {
        New-Item $errPath -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Script.errs | Export-Clixml $(Join-Path $errPath "errs.xml") -Depth 5
}

$allDone = $true
} # end of try for transcript
finally {
    Remove-Runspace

    if (-not $allDone) {
        Write-Log "Script was interrupted in the middle of execution."
    }

    Write-Log "Total time is $(((Get-Date) - $startDateTime).TotalSeconds) seconds"
    Close-Log

    # release transcript file even when script is stopped in the middle.
    if ($transcriptEnabled) {
        $(Stop-Transcript) 2>&1 | Out-Null
    }
}

$zipFileName = "$($OrgName)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

if ($SkipZip) {
    # PSv2 does not have LiteralPath parameter.
    Rename-Item -Path $Path -NewName $zipFileName
}
else {
    Write-Progress -Activity $collectionActivity -Status:"Packing into a zip file" -PercentComplete:95
    Compress-Folder -Path:$Path -ZipFileName:$zipFileName -Destination:$originalPath | Out-Null

    $err = $(Remove-Item $Path -Force -Recurse) 2>&1
    if ($err) {
        Write-Warning "Failed to delete a temporary folder `"$Path`". $err"
    }
}

Write-Progress -Activity $collectionActivity -Status "Done" -Completed
if ($allDone) {
    Write-Host "Done!" -ForegroundColor Green
}
