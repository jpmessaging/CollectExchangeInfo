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

.PARAMETER KeepOutputFiles
Switch to keep the output files. If this is not specified, all the output files will be deleted after being packed to a zip file.
In order to avoid deleting unrelated files or folders, this script makes sure that the folder specified by Path paramter is empty and if not empty, it stops executing.

.EXAMPLE
.\Collect-ExchangeInfo -Path .\exinfo -Servers:* 

Create (if not exist) a sub folder "exinfo" under the current path.
All the output files are saved in this folder.
All Exchange servers will be accessed since * is specified for "Servers".

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
param
(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [string[]]$Parameters = @(),    
    [string[]]$Servers = @(),
    [switch]$IncludeFIPS,    
    [switch]$IncludeEventLogs = $false,    
    [switch]$IncludeEventLogsWithCrimson,
    [switch]$IncludeIISVirtualDirectories,
    [switch]$KeepOutputFiles
)

$version = "2019-08-08"
#requires -Version 2.0

<#
  Save object(s) to a text file and optionally export to CliXml.
  When a string is given, it's assumed to be an error and it's saved in the file specified as ErrorPath
#>
function Save-Object
{
    [CmdletBinding()]
    Param(
        #[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Parameter(ValueFromPipeline=$true)]
        $object,
        $Name,
        [string]$Path = $Script:Path,
        [string]$ErrorPath = [System.IO.Path]::Combine($Path, "errors.txt"),
        [bool]$WithCliXml = $true,
        $Depth = 5 # depth for Export-CliXml
    )
 
    BEGIN
    {     
        # Need to accumulate result to support pipeline. Use List<> to improve performance
        $objectList = New-Object System.Collections.Generic.List[object]
        [string]$objectName = $Name
    }
    
    PROCESS
    {
        # Validate the given objects.  If valid, collect them in a list.
        # Collected objects are outputted in the END block
        
        # When explicitly passed, object is actually a list of objects.
        # When passed from pipeline, object is a single object.
        # To deal with this, use foreach.
        
        foreach ($o in $object)
        {
            if ($o -eq $null)
            {
                return
            }
            <#
            elseif($o -is [string])
            {                
                # assume a string object is an error and write it to log
                Write-Log $o
            }
            #>
            else 
            {                
                if (-not($objectName))
                {
                    $objectName = $o.GetType().Name            
                }
                $objectList.Add($o)
            }
        }   
    }
    
    END
    {
        if ($objectList.Count -gt 0)
        {
            if(-not $objectName)
            {
                Write-Log "[Save-Object] Error:objectName is null"
            }

            if ($WithCliXml)
            {                
                $objectList | Export-Clixml -Path:([System.IO.Path]::Combine($Path, "$objectName.xml")) -Encoding:UTF8 -Depth $Depth
            }
            
            $objectList | select * | Out-File ([System.IO.Path]::Combine($Path, "$objectName.txt")) -Encoding:UTF8            
        }
    }
}
 
<#
  Zip a folder
  Path is the folder to zip
  ZipFileName is the name of the zip file to create. may or may not have .zip extention 
#>
function Zip-Folder
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]$Path,
        [string]$ZipFileName,
        [bool]$IncludeDateTime = $true,
        [switch]$RemoveFiles
    )
     
    $Path = Resolve-Path $Path
     
    $zipFileNameWithouExt = [System.IO.Path]::GetFileNameWithoutExtension($ZipFileName)
    if ($IncludeDateTime)
    {
        # Create a zip file in TEMP folder with current date time in the name
        # e.g. Contoso_20160521_193455.zip
        $currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
        $zipFileName = $zipFileNameWithouExt + "_" + "$currentDateTime.zip"                
    }
    else
    {
        $zipFileName = "$zipFileNameWithouExt.zip"
    }
    $zipFilePath = Join-Path ((Get-Item ($env:TEMP)).FullName) -ChildPath $zipFileName

    $NETFileSystemAvailable = $true
 
    try
    {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    }
    catch 
    {
        Write-Warning "System.IO.Compression.FileSystem wasn't found. Using alternate method"
        $NETFileSystemAvailable = $false
    }
 
    if ($NETFileSystemAvailable)
    {
        [System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $zipFilePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
    }
    else 
    {
        # Use Shell.Application COM
        $delayMilliseconds = 200
 
        # Create a zip file manually
        $shellApp = New-Object -ComObject Shell.Application        
        Set-Content $zipFilePath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (Get-Item $zipFilePath).IsReadOnly = $false
 
        $zipFile = $shellApp.NameSpace($zipFilePath)
  
        $items = Get-ChildItem -Path:"$Path"
        $completedItemCount = 0        
    
        # Idea1: copy a whole directory: 
        # Better throughput overall
        # no item-wise progress
 
        # Start copying the whole and wait until it's done. Note: CopyHere works asynchronously.
        $zipFile.CopyHere($Path)
 
        # Now wait
        $inProgress = $true
        Sleep -Milliseconds 3000
        [System.IO.FileStream]$file = $null
        while ($inProgress)
        {
            Sleep -Milliseconds $delayMilliseconds
            
            try
            {
                $file = [System.IO.File]::Open($zipFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                $inProgress = $false
            }
            catch [System.IO.IOException]
            {
                Write-Debug $_.Exception.Message
            }
            finally
            {
                if ($file -ne $null)
                {
                    $file.Close()
                }
            }
        }
    }
            
    # Move the zip file from TEMP folder to Path
    if (Test-Path $zipFilePath)
    {
        Move-Item $zipFilePath -Destination $Path

        # If requested, remove zipped files 
        if ($RemoveFiles)
        {
            # At this point, don't use Write-Log since the log file will be deleted too
            Write-Verbose "Removing zipped files"
            Get-ChildItem $Path -Exclude $ZipFileName | Remove-Item -Recurse -Force 
        }
    }
    else 
    {
        Write-Output "Zip file wasn't successfully created at $zipFilePath"        
    }    
}
 
<#
  Runs Ldifde for Exchange organization in configuration context
  Note: Path must exist; otherwise it just returns error string
#>
function Run-Ldifde
{
    [Parameter(Mandatory=$true)]
    param
    (
      [string]$Path,
      [string]$FileName = "Ldifde.txt"
    )
        
    # if Path doesn't exit, create it
    if (!(Test-Path $Path))
    {
       New-item -ItemType directory $Path | Out-Null
    } 
    $resolvedPath  = Resolve-Path $Path -ErrorAction SilentlyContinue
    $filePath = Join-Path -Path $resolvedPath -ChildPath $FileName
 
    # Check if Ldifde.exe exists
    $IsLdifdeAvailable = $false
    foreach ($path in $env:Path.Split(";"))
    {        
        if ($path)
        {
            $exePath = Join-Path -Path $Path -ChildPath "ldifde.exe"
            if (Test-Path $exePath)
            {
                $IsLdifdeAvailable = $true;
                break;
            }
        }
    }
    if (!$IsLdifdeAvailable)
    {        
        return "[Run-Ldifde] Ldifde is not available"        
    }
    
    $exorg = (Get-OrganizationConfig).DistinguishedName
    
    if (!$exorg)
    {
        return "[Run-Ldifde] Couldn't get Exchange org DN"
    }
    
    # If this is an Edge server, use a port 50389.
    $server = Get-ExchangeServer $env:COMPUTERNAME
    if ($server -and $server.IsEdgeServer)
    {
        $Port = 50389
    }

    try
    {
        $fileNameWihtoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)        
        $stdOutput = Join-Path $resolvedPath -ChildPath "$fileNameWihtoutExtension.out"
        
        if ($Port)
        {
    		$process = Start-Process ldifde -ArgumentList "-u -d `"$exorg`" -s localhost -t $Port -f `"$filePath`"" -PassThru -NoNewWindow -RedirectStandardOutput:$stdOutput
        }
        else
        {
            $process = Start-Process ldifde -ArgumentList "-u -d `"$exorg`" -f `"$filePath`"" -PassThru -NoNewWindow -RedirectStandardOutput:$stdOutput
        }
        if (!$process.HasExited)
        {
            Wait-Process -InputObject $process         
        }        
        
        $process = $null        
    }
    finally
    {
        if ($process)
        {
            Stop-Process -InputObject:$process
            Write-Output "[Run-Ldifde] ldifde cancelled"
        }
    }
}

<#
  Run a given command only if it's available
  Run with parameters specified as Global Parameter (i.e. $script:Parameters)
#>
function RunCommand
{
    [CmdletBinding()]    
    param(
    [Parameter(Mandatory=$true)]
    [string]$Command
    )
    
    $endOfCmdlet = $Command.IndexOf(" ")
    if ($endOfCmdlet -lt 0)
    {
        $cmdlet = $Command
    }
    else 
    {
        $cmdlet = $Command.Substring(0, $endOfCmdlet)
    }
 
    # check if cmdlet is available
    $cmd = Get-Command $cmdlet -ErrorAction:SilentlyContinue
    if ($cmd -eq $null)
    {
        Write-Log "$cmdlet is not available"
        return
    } 

    # check params
    # if any explicitly-requested params are not available, bail    
    $paramMatches = Select-String " -(?<paramName>\w+)" -Input $Command -AllMatches

    if ($paramMatches)
    {
        $params = @(
        foreach($paramMatch in $paramMatches.Matches)
        {
            $paramName = $paramMatch.Groups['paramName'].Value  

            # In order to support non-exact match, check each key
            $keyMatch = @(
            foreach ($key in $cmd.Parameters.keys)  
            {
                if ($key -like "$($paramName)*")
                {
                    $key
                }
            }
            )

            # if there's no match or too many matches, bail.
            if ($keyMatch.Count -eq 0)
            {
                Write-Log "Parameter '$paramName' is not available for $cmdlet"
                return 
            }             
            elseif ($keyMatch.Count -gt 1)
            {
                Write-Log "Parameter '$paramName' is ambiguous for $cmdlet"
                return 
            }

            $keyMatch[0]
        }
        )
    }

    # check if any parameter is requested globally
    # it's ok if these parameters are not available.
    foreach ($param in $script:Parameters)
    {
        $paramName = ($param -split ":")[0]

        if ($cmd.Parameters[$paramName] -ne $null)
        {
            # explicitly-requested params take precedence
            # if not already in the list, add it.
            if ($params -notcontains $paramName)
            {
                $Command += " -$param"
           }
        }
    }

    # Finally run the command
    Write-Log "Running $Command"
    try
    {
        $err = $($o = Invoke-Expression $Command) 2>&1
        if ($err)
        {
            Write-Log "[Error] `"$Command`" failed. $err"
        }
        else
        {
            Write-Output $o
        }
    }
    catch 
    {
        Write-Log "[Error] `"$Command`" failed. $_"
    }

    <#
    try
    {
        Invoke-Expression $Command
    }
    catch 
    {
        # Log error and continue
        Write-Log $_
    }
    #>
}
 
<#
  Run command against servers
#>
function Run
{
    [CmdletBinding()]    
    param(
    [Parameter(Mandatory=$true)]
    [string]$Command,    
    [string[]]$Servers,
    [string]$Identifier = "Server",
    [bool]$ExecuteWhenNoServers = $true,
    [Parameter(ValueFromPipeline=$true)]
    [object[]]$ResultCollection,
    [bool]$AllowDuplicate = $true,
    [switch]$PassThru
    )

    BEGIN
    {
        $result = New-Object System.Collections.Generic.List[object]
    }
    # Accumulate the previous results
    PROCESS
    {
        # Make sure not to add $null and collection itself
        # $ResultCollection | where {$_ -ne $null} | ForEach {$result.Add($_)}
        foreach ($pipedObj in $ResultCollection)
        {                        
            # In PowerShellV2, $null is iterated over.
            if ($pipedObj)
            {
                $result.Add($pipedObj)
            }
        }
    }

    END
    {
        if ($ResultCollection -and ($ResultCollection.Count -ge $allExchangeServers.Count))
        {            
            Write-Log "Pipeline input has already $($allExchangeServers.Count) objects. Skipping `"$Command`""                        
        }
        elseif (!$Servers -and $ExecuteWhenNoServers)
        {
            foreach ($o in @(RunCommand $Command))
            {
                # Check duplicates
                if (-not $AllowDuplicate)
                {
                    $dups = @($result | where {$_.Distinguishedname -eq $o.Distinguishedname})
                    if ($dups.Count -gt 0)
                    {
                        # this is a duplicate. skip.                     
                        Write-Log "`"dropping a duplicate: '$($o.Distinguishedname)'`""
                        continue
                    }
                }
                $result.Add($o)
            }
        }
        elseif ($Servers)
        {
            foreach ($server in $Servers)
            {       
                $firstTimeAddingServerName = $true
                foreach ($entry in @(RunCommand "$Command -$Identifier $server"))
                {
                    # Add ServerName prop if not exist already (but log only the first time per cmdlet)                    
                    if (!$entry.ServerName -and !$entry.Server -and !$entry.ServerFqdn -and !$entry.MailboxServer -and !$entry.Fqdn)
                    {
                        if ($firstTimeAddingServerName)
                        {
                            Write-Log "Adding ServerName to the result of '$Command -$Identifier $server'"
                            $firstTimeAddingServerName = $false
                        }

                        # This is for PowerShell V2
                        # $entry | Add-Member -Type NoteProperty -Name:ServerName -Value:$server
                        $entry = $entry | select *, @{N='ServerName';E={$server}}
                    }

                    $result.Add($entry)                      
                }                
            }
        }

        if ($PassThru)
        {
            Write-Output $result
        }
        else
        {   
            # Extract cmdlet name (e..g "Get-MailboxDatabase" -> "MailboxDatabase")        
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
function Write-Log
{
    [CmdletBinding()] 
    param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]$Text,
    [string]$Path = $Script:logPath
    )
 
    $currentTime = get-date
    $currentTimeFormatted = $currentTime.ToString("yyyy/MM/dd HH:mm:ss.fffffff(K)")
    
    [System.TimeSpan]$delta = 0;
    if ($Script:lastLogTime -ne $null)
    {
        $delta = $currentTime.Subtract($Script:lastLogTime)
    }
    else
    {
        # For the first time, add header
        Add-Content $Path "date-time,delta(ms),info"       
    }
    
    Write-Verbose $Text     
    Add-Content $Path "$currentTimeFormatted,$($delta.TotalMilliseconds),$text"    
    $Script:lastLogTime = $currentTime
}
 
<#
  Run Get-*VirtualDirectory & Get-OutlookAnywhere for all servers in $Servers
  If IncludeIISVirtualDirectories is specified, access IIS vdir for servers == IsDirectAccess.  
  Otherwise, only AD info will be collected
#>
function Get-VirtualDirectories
{
    [CmdletBinding()]    
    param()
    # List of Get-*VirtualDirectory commands.
    # CommantType can be different depending on whether Local PowerShell or Remote PowerShell
    $commands = Get-Command Get-*VirtualDirectory -ErrorAction:SilentlyContinue | where {$_.name -ne 'Get-WebVirtualDirectory'}
    $commands += Get-Command Get-OutlookAnywhere -ErrorAction:SilentlyContinue

    foreach ($command in $commands)
    {             
        # If ShowMailboxVirtualDirectories param is available, add it (E2013 & E2016).         
        if ($command.Parameters -and $command.Parameters.ContainsKey('ShowMailboxVirtualDirectories'))
        {
            # if IncludeIISVirtualDirectories, then access direct access servers. otherwise, don't touch servers (only AD)
            if ($IncludeIISVirtualDirectories)
            {
                Run "$($command.Name) -ShowMailboxVirtualDirectories" -Servers:($allExchangeServers | where {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -ExecuteWhenNoServers:$false -PassThru |
                    Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -AllowDuplicate:$false            
            }
            else
            {
                Run "$($command.Name) -ADPropertiesOnly -ShowMailboxVirtualDirectories" -AllowDuplicate:$false                                  
            }
        }
        else
        {
            if ($IncludeIISVirtualDirectories)
            {   
                Run "$($command.Name)" -Servers:($allExchangeServers | where {$_.IsExchange2007OrLater -and $_.IsClientAccessServer -and $_.IsDirectAccess}) -ExecuteWhenNoServers:$false -PassThru |
                    Run "$($command.Name) -ADPropertiesOnly" -AllowDuplicate:$false
            }
            else
            {
                Run "$($command.Name) -ADPropertiesOnly" -AllowDuplicate:$false
            }
        }
    }
}

function Run-FIPSCmdlet
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Servers,
        [Parameter(Mandatory = $true)]
        [string]$FIPSCmdlet    
    )
  
    $result = @(
    foreach($server in $Servers)
    {
        $command = "Add-PSSnapin -Name Microsoft.Forefront.Filtering.Management.PowerShell;"
        $command += "$FIPSCmdlet;"
        $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($command) 
            
        Write-Log "[FIPS] Running $FIPSCmdlet at $server"
        Invoke-Command -ComputerName $server -ScriptBlock $scriptblock -ErrorAction SilentlyContinue   
    }
    )
    $commandName = $FIPSCmdlet.Substring(4)
    $result | Save-Object -Name $commandName
}

function Run-FIPSCommands
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Servers
    )

    PROCESS
    {
        $command = "Add-PSSnapin -Name Microsoft.Forefront.Filtering.Management.PowerShell;"
        $command += "Get-Command -Module Microsoft.Forefront.Filtering.Management.PowerShell"            
        $scriptblock = $ExecutionContext.InvokeCommand.NewScriptBlock($command) 

        # if no server is given, bail
        if ($Servers -eq $null)
        {
            return
        }

        # ASSUME all servers have the same FIPS cmdlets        
        $FIPSCmdlets = Invoke-Command -ComputerName:$Servers[0] -ScriptBlock $scriptblock -ErrorAction SilentlyContinue   
        # filter only Get-* cmdlets except Get-ConfigurationValue
        $FIPSCmdlets = $FIPSCmdlets | where {$_.Name -like "Get-*" -and $_.Name -ne "Get-ConfigurationValue" }

        foreach ($cmdlet in $FIPSCmdlets)
        {
            Run-FIPSCmdlet -Servers:$Servers -FIPSCmdlet:$cmdlet
        }
    }
}

function Get-SPN 
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        $Path # folder path to save output. 
    )  
          
    # Make sure Path exists; if not, just return error string
    $resolvedPath  = Resolve-Path $Path -ErrorAction SilentlyContinue
    if (!$resolvedPath)
    {
        return "[Get-SPN] $Path doesn't exist"
    }
 
    $filePath = Join-Path -Path $Path -ChildPath "setspn.txt"

    # Check if setspn.exe exists
    $isSetSPNAvailable = $false
    foreach ($path in $env:Path.Split(";"))
    {        
        if ($path)
        {
            $exePath = Join-Path -Path $Path -ChildPath "setspn.exe"
            if (Test-Path $exePath)
            {
                $isSetSPNAvailable = $true;
                break;
            }
        }
    }
    if (!$isSetSPNAvailable)
    {        
        $msg = "setspn is not available"               
        Write-Log $msg
        return
    }
    
    Add-Content -Path:$filePath -Value:"[setspn -P -F -Q http/*]"            
    $result = Run-ShellCommand -FileName setspn -Argument '-P -F -Q http/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeMDB/*]"    
    $result = Run-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeMDB/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath         

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeRFR/*]"    
    $result = Run-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeRFR/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath   

    Add-Content -Path:$filePath -Value:"$([Environment]::NewLine)[setspn -P -F -Q exchangeAB/*]"
    $result = Run-ShellCommand -FileName setspn -Argument '-P -F -Q exchangeAB/*' -Wait
    $result.StdOut | Add-Content -Path:$filePath   
}

function Run-ShellCommand
{
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

    if (!$Wait)
    {        
        Write-Output $process
    }
    else
    {
        # deadlock can occur b/w parent and child proces!
        # https://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandardoutput(v=vs.110).aspx

        #$process.BeginErrorReadLine()
        $stdout = $process.StandardOutput.ReadToEnd()        
        $process.WaitForExit()
        
        $result = New-Object -TypeName PSCustomObject -Property @{Process = $process; StdOut = $stdout; ExitCode = $exitCode}
        Write-Output $result
    }
}

function Get-MSInfo32
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        $Servers
    )  

    try
    {
        foreach ($server in $Servers)
        {
            Write-Log "Running $($MyInvocation.MyCommand) on $server"

            $nfoFilePath = Join-Path $Script:Path -ChildPath "$server.nfo"                
            $process = Start-Process "msinfo32.exe" -ArgumentList "/computer $server /nfo $nfoFilePath" -PassThru
            if (Get-Process -Id:($process.Id) -ErrorAction:SilentlyContinue)
            {
                Wait-Process -InputObject:$process
            }
    }
    }
    finally
    {
        if ($process -and (Get-Process -Id:($process.Id) -ErrorAction:SilentlyContinue))
        {
            Write-Log "[$($MyInvocation.MyCommand)] msinfo32 cancelled for $server"
            Stop-Process -InputObject $process
        }
    }
}

function Collect-EventLogs
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]  
        $Path,
        $Computers = @($env:COMPUTERNAME),
        [switch]$IncludeCrimsonLogs,
        [switch]$Zip
    )

    if (!(Test-Path $Path))
    {
        New-item -ItemType directory $Path | Out-Null
    }

    foreach ($computer in $computers)
    {
        # By default, collect app ans sys logs
        $logs = "Application","System"

        $savePath = Join-Path $Path -ChildPath $computer
        # Create a folder for each computer
        if (!(Test-Path $savePath))
        {
            New-item -ItemType directory $savePath  | Out-Null
        }        

        Write-Log "[$($MyInvocation.MyCommand)] Saving event logs on $computer ..."

        # Detect machine-local Window's TEMP path (i.e. C:\Windows\Temp)
        # Logs are saved here temporarily and will be moved to Path
        $win32os = Get-WmiObject win32_operatingsystem -ComputerName:$computer
        if (!$win32os)
        {
            # WMI failed. Maybe wrong server name?
            Write-Log "[$($MyInvocation.MyCommand)] Get-WmiObject win32_operatingsystem failed for '$computer'"
            continue
        }

        # This is remote machine's path
        $currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
        $localPath = Join-Path $win32os.WindowsDirectory -ChildPath "Temp\EventLogs_$currentDateTime"
        $uncPath = "\\$computer\" + $localPath.Replace(':','$') 
        if (!(Test-Path $uncPath))
        {
            New-item -ItemType directory $uncPath | Out-Null
        }

        # For Crimson logs
        if ($IncludeCrimsonLogs)
        {
            $logs += (wevtutil el /r:$computer) -like "Microsoft-Exchange*" 
        }

        foreach ($log in $logs)
        {
            Write-Log "[$($MyInvocation.MyCommand)] Saving $log ..."
            $fileName = $log.Replace('/', '_') + '.evtx' 
            $localFilePath = Join-Path $localPath -ChildPath $fileName
            $uncFilePath = "\\$computer\" + $localFilePath.Replace(':','$')    

            wevtutil epl $log $localFilePath /ow /r:$computer 
        }

        # Try to zip up before copying in order to save bandwidth.
        # This is possible only if remote management is enabled on the remote machine (i.e. winrm quickconfig)
        $zipFileName = "EventLogs_$computer.zip"
        $zipCreated = $true
        try
        {
            Invoke-Command -ComputerName $computer -ScriptBlock ${function:Zip-Folder} -ArgumentList $localPath,$zipFileName,$false -ErrorAction Stop
        }
        catch 
        {
            Write-Log "[$($MyInvocation.MyCommand)] Cannot create a zip file on $computer. Each event log file will be copied."
            $zipCreated = $false
        }

        if ($zipCreated)
        {
            Write-Log "[$($MyInvocation.MyCommand)] Copying a zip file '$zipFileName' from $computer"
            Move-Item (Join-Path $uncPath -ChildPath $zipFileName) -Destination $savePath -Force
        }
        else
        {
            Write-Log "[$($MyInvocation.MyCommand)] Copying *evtx files from $computer"
            $evtxFiles = Get-ChildItem -Path $uncPath -Filter '*.evtx'
            foreach ($file in $evtxFiles)
            {
                Move-Item $file.FullName -Destination $savePath -Force
            }
        }

        # Clean up
        Remove-Item $uncPath -Recurse
     }
}

function Get-DAG
{    
    $dags = RunCommand Get-DatabaseAvailabilityGroup
    $result = @(
    foreach ($dag in $dags)
    {
        # Get-DatabaseAvailabilityGroup with "-Status" fails for cross Exchange versions (e.g. b/w E2010, E2013)
        $dagWithStatus = RunCommand "Get-DatabaseAvailabilityGroup $dag -Status -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue"
        if ($dagWithStatus)
        {            
            $dagWithStatus
        }
        else 
        {
            Write-Log "[$($MyInvocation.MyCommand)] Get-DatabaseAvailabilityGroup $($dag.Name) -Status failed. The result without -Status will be saved."
            $dag
        }
    }
    )

    Save-Object $result -Name "DatabaseAvailabilityGroup"
}

function Get-DotNetVersion
{
    [CmdletBinding()]
    param (
        $Computer = $env:COMPUTERNAME
    )    

    # Read NDP registry
    try
    {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
    }
    catch
    {
        Write-Error "Couldn't open registry key of $Computer.`n$_"
        return
    }
   
    $ndpKey = $reg.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP")
    $result = @(
    foreach ($versionKeyName in $ndpKey.GetSubKeyNames())
    {
        # ignore "CDF" etc
        if ($versionKeyName -notlike "v*") 
        {
            continue
        }

        $versionKey = $ndpKey.OpenSubKey($versionKeyName)
        $version = $versionKey.GetValue("Version", "")
        $sp = $versionKey.GetValue("SP", "")
        $install = $versionKey.GetValue("Install", "")
        
        if ($version)
        {
            New-Object PSCustomObject -Property @{Version = $version; SP = $sp; Install = $install; SubKey = $null; Release = $release; NET45Version = $null; ServerName = $Computer}           
            continue
        }

        # for v4 and V4.0, check sub keys            
        foreach ($subKeyName in $versionKey.GetSubKeyNames())
        {
            $subKey = $versionKey.OpenSubKey($subKeyName)
            $version = $subKey.GetValue("Version", "")
            $install = $subKey.GetValue("Install", "")
            $release = $subKey.GetValue("Release", "")
            if ($release)
            {
                $NET45Version = Get-Net45Version $release
            }
            else
            {
                $NET45Version = $null
            }
            New-Object PSCustomObject -Property @{Version = $version; SP = $sp; Install = $install; SubKey = $subKeyName;Release = $release; NET45Version = $NET45Version; ServerName = $Computer}
        }
    }
    )

    $result = $result | sort -Property Version
    Write-Output $result
}

function Get-Net45Version    
{
    [CmdletBinding()]
    param (
    [Parameter(Mandatory=$True)]
    $Release
    )
     
    $version = $null;

    switch ($Release)
    {
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


function Get-TlsRegistry
{
    [CmdletBinding()]
    param(
        $Computer= $env:COMPUTERNAME
    )

    try
    {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
    }
    catch
    {
        Write-Error "Couldn't open registry key of $Computer.`n$_"
        return
    }

    $result = New-Object System.Collections.Generic.List[object]

    # OS SChannel related   
    $protocols = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\')
    # "Protocols" key should exist
    foreach ($protocolKeyName in $protocols.GetSubKeyNames())
    {
        # subKeyName is "SSL 2.0", "TLS 1.0", etc
        $protocolKey = $protocols.OpenSubKey($protocolKeyName)
        #$subkeyNames = @('Client','Server')
        foreach ($subKeyName in $protocolKey.GetSubKeyNames())
        {
            $subKey = $protocolKey.OpenSubKey($subKeyName)

            $disabledByDefault = $subKey.GetValue('DisabledByDefault', '')
            $enabled = $subKey.GetValue('Enabled', '')        
            
            $result.Add((New-Object PSCustomObject -Property @{ServerName = $Computer
                Name="SChannel $protocolKeyName $subKeyName"
                DisabledByDefault = $disabledByDefault
                Enabled = $enabled
                RegistryKey = $subKey.Name
                })
            )
        }
    }

    # .NET related
    $netKeyNames = @('SOFTWARE\Microsoft\.NETFramework\', 'SOFTWARE\Wow6432Node\Microsoft\.NETFramework\')

    foreach ($netKeyName in $netKeyNames)
    {
        $netKey = $reg.OpenSubKey($netKeyName)
        $netSubKeyNames = @('v2.0.50727','v4.0.30319')

        foreach ($subKeyName in $netSubKeyNames)
        {
            $subKey = $netKey.OpenSubKey($subKeyName)
            if (-not $subKey)
            {
                continue
            }

            $systemDefaultTlsVersions = $subKey.GetValue('SystemDefaultTlsVersions','')
            $schUseStrongCrypto = $subKey.GetValue('SchUseStrongCrypto','')

            if ($subKey.Name.IndexOf('Wow6432Node', [StringComparison]::OrdinalIgnoreCase) -ge 0)
            {
                $name = ".NET Framework $subKeyName (Wow6432Node)"
            }
            else
            {
                $name = ".NET Framework $subKeyName"
            }

            $result.Add((New-Object PSCustomObject -Property @{ServerName = $Computer
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

function Get-TCPIP6Registry
{
    [CmdletBinding()]
    param(
        $Computer= $env:COMPUTERNAME
    )

    try
    {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer)
    }
    catch
    {
        Write-Error "Couldn't open registry key of $Computer.`n$_"
        return
    }

    $key = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\') 
    $disabledComponents = $key.GetValue('DisabledComponents','')

    New-Object PSCustomObject -Propert @{DisabledComponents = $disabledComponents}
}

function Get-IISWebBinding
{
    [CmdletBinding()]
    param(
    $Server
    )
    
    $block = {       
        $err = Import-Module WebAdministration 2>&1
        if ($err)
        {
            Write-Error "Import-Module WebAdministration failed."
        }
        else
        {
            Get-WebBinding
        }
    }

    if ($Server -eq $env:COMPUTERNAME)
    {        
        Invoke-Command -ScriptBlock $block
        return
    }
    
    $sess = New-PSSession -ComputerName $server -ErrorAction SilentlyContinue 
    if (-not $sess)
    {
        Write-Error "Failed to create a remote session to $server."
        return
    }
        
    Invoke-Command -Session $sess -ScriptBlock $block
    Remove-PSSession $sess
}

<#
  Main
#> 

# If the path already exists, make sure it's empty in case $KeepOutputFiles is not specified.
# otherwise, those files will be deleted after packed into a zipped file.
# If the path doesn't exist, create it.
if (Test-Path $Path)
{
    if (!$KeepOutputFiles)
    {
        if (@(Get-ChildItem $Path).Count -ne 0)
        {
            # to be safe, bail
            Write-Warning "File(s) or folder(s) exist in $Path. Please use a different path or add KeepOutputFiles switch"
            return
        }
    }
}
else
{    
    New-item -ItemType directory $Path | Out-Null 
}

# Resolve Path in case a relative path is given
$Path = Resolve-Path $Path

# Prepare for logging 
# NOTE: until $logPath is defined, don't call Write-Log
$logFileName = "Log.txt" 
$logPath = Join-Path -Path $Path -ChildPath $logFileName

$cmd = Get-Command "Get-OrganizationConfig" -ErrorAction:SilentlyContinue
if ($cmd -eq $null)
{
    Write-Error "Get-OrganizationConfig is not available. Please run with Exchange Remote PowerShell session"
    Write-Log "Get-OrganizationConfig is not available. Please run with Exchange Remote PowerShell session"
    return
}
$OrgName = (Get-OrganizationConfig -WarningAction:SilentlyContinue).Identity

$lastLogTime = $null
Write-Log "Organization Name = $OrgName"
Write-Log "Script Version=$version"
Write-Log "COMPUTERNAME=$env:COMPUTERNAME"

# Log parameters (raw values are in $PSBoundParameters, but want fixed-up values (e.g. Path)
$sb = New-Object System.Text.StringBuilder
foreach ($paramName in $PSBoundParameters.Keys)
{
    $var = Get-Variable $paramName -ErrorAction SilentlyContinue
    if ($var)
    {
        $sb.Append("$($var.Name):$($var.Value); ") | Out-Null
    }
}
Write-Log $sb.ToString()


# Prepare the list of Exchange Servers to directly access by parsing the values specified in "Servers" parameter
# Used in VDir, Mailbox Catabase Copy, Certificate etc.
# First, get the candidates from the user specified values in $Servers
$directAccessCandidates =@(
foreach ($server in $Servers)
{    
    # $server's value might be something like "e2013*" and matches multiple servers
    $exServers = @(Get-ExchangeServer $server)

    if ($exServers.Count -eq 0)
    {
        Write-Log "Get-ExchangeServer did not find any server matching '$server'"
    }

    foreach ($exServer in $exServers)
    {        
        # In PowerShellv2. $exServer may be $null.
        if ($exServer -eq $null)
        {
            continue
        }
        
        # add if it's not a duplicate
        $inDAS = ($directAccessCandidates | where {$_.Name -eq $exServer.Name}) -ne $null
        if (!$inDAS)
        {
            $exServer         
        }
    }
}
)

Write-Log "directAccessCandidates = $directAccessCandidates"

# Now test connectivity to those servers
# Since there shouldn't be anything blocking communication b/w Exchange servers, we should be able to use ICMP
#[Microsoft.Exchange.Data.Directory.Management.ExchangeServer[]]$directAccessServers = @()
$directAccessServers = @(
foreach ($server in $directAccessCandidates)
{
    if (Test-Connection -ComputerName:$server.Name -Count 1 -Quiet)
    {        
        $server
    }
    else
    {
        Write-Log "Connectivity test failed on $server"
    }
}
)
Write-Log "directAccessServers = $directAccessServers"

$allExchangeServers = @(Get-ExchangeServer)
$allExchangeServers | Add-Member -Type NoteProperty -Name IsDirectAccess -Value:$false

foreach ($server in $allExchangeServers)
{
    if (($directAccessServers | where {$_.Name -eq $server}) -ne $null)
    {
        $server.IsDirectAccess = $true
    }
}

#
# Start collecting
#
$transcriptPath = Join-Path -Path $Path -ChildPath "transcript.txt"
$transcriptEnabled = $false
try
{
    Start-Transcript -Path $transcriptPath -NoClobber -ErrorAction:Stop
    $transcriptEnabled = $true
}
catch 
{    
    Write-Log "Start-Transcript is not available"
}

# Start of try for transcript
try
{
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
Run Get-IPAllowListEntry -Servers:($directAccessServers | where {$_.IsE14OrLater -and $_.IsHubTransportServer}) -ExecuteWhenNoServers:$false
Run Get-IPAllowListProvider 
Run Get-IPAllowListProvidersConfig 
Run Get-IPBlockListConfig 
Run Get-IPBlockListEntry -Servers:($directAccessServers | where {$_.IsE14OrLater -and $_.IsHubTransportServer}) -ExecuteWhenNoServers:$false
Run Get-IPBlockListProvider 
Run Get-IPBlockListProvidersConfig 
Run Get-JournalRule 
Run Get-RecipientFilterConfig 
Run Get-RMSTemplate 
Run Get-SenderFilterConfig 
Run Get-SenderIdConfig 
Run Get-SenderReputationConfig 
Run Get-TransportRule 
# these cmdlets are meant to run locally and don't have server specifiers (-Server, -Identity)
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
Write-Log "AD & AddressBook & Retention Done"
 
# Server Settings
Write-Progress -Activity $collectionActivity -Status:"Server Settings" -PercentComplete:40
 
Run Get-ExchangeServer      
Run Get-MailboxServer

# For CAS (>= E14) in DAS list, include ASA info
Run "Get-ClientAccessServer -IncludeAlternateServiceAccountCredentialStatus -WarningAction:SilentlyContinue" -Servers:($allExchangeServers | where {$_.IsDirectAccess -and $_.IsClientAccessServer -and -$_.IsE14OrLater}) -Identifier:Identity -ExecuteWhenNoServers:$false -PassThru | 
    Run "Get-ClientAccessServer -WarningAction:SilentlyContinue" -Identifier:Identity -AllowDuplicate:$false

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

Run "Get-MailboxDatabase -Status -IncludePreExchange" -Servers:($allExchangeServers | where {$_.IsMailboxServer -and $_.IsDirectAccess}) -ExecuteWhenNoServers:$false -PassThru |
    Run "Get-MailboxDatabase -IncludePreExchange" -AllowDuplicate:$false

Run "Get-PublicFolderDatabase -Status" -Servers:($allExchangeServers | where {$_.IsMailboxServer -and $_.IsDirectAccess}) -ExecuteWhenNoServers:$false -PassThru | 
    Run "Get-PublicFolderDatabase" -AllowDuplicate:$false

Run Get-MailboxDatabaseCopyStatus -Servers:($directAccessServers | where {$_.IsE14OrLater -and $_.IsMailboxServer}) -ExecuteWhenNoServers:$false
Get-DAG
Run Get-DatabaseAvailabilityGroupConfiguration
if (Get-Command Get-DatabaseAvailabilityGroup -ErrorAction:SilentlyContinue)
{
    Run "Get-DatabaseAvailabilityGroupNetwork -ErrorAction:SilentlyContinue" -Servers:(Get-DatabaseAvailabilityGroup) -Identifier:'Identity'
}
Write-Log "Database Done"
 
# Virtual Directories
Write-Progress -Activity $collectionActivity -Status:"Virtual Directory Settings" -PercentComplete:60
Get-VirtualDirectories
Run "Get-IISWebBinding" -Servers $directAccessServers -ExecuteWhenNoServers:$false -PassThru | Save-Object -Name WebBinding

# Active Monitoring & Managed Availability
Write-Progress -Activity $collectionActivity -Status:"Monitoring Settings" -PercentComplete:70
Run Get-GlobalMonitoringOverride 
Run Get-ServerMonitoringOverride -Servers:($directAccessServers | where {$_.IsE15OrLater})  -ExecuteWhenNoServers:$false
Run Get-ServerComponentState -Servers:($directAccessServers | where {$_.IsE15OrLater}) -Identifier:Identity -ExecuteWhenNoServers:$false
Run Get-HealthReport -Servers:($directAccessServers | where {$_.IsE15OrLater}) -Identifier:Identity -ExecuteWhenNoServers:$false
Run Get-ServerHealth -Servers:($directAccessServers | where {$_.IsE15OrLater}) -Identifier:Identity -ExecuteWhenNoServers:$false
Run Test-ServiceHealth -Servers:$directAccessServers -ExecuteWhenNoServers:$false

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
Run Get-ExchangeCertificate -Servers:($directAccessServers | where {$_.IsE14OrLater}) -ExecuteWhenNoServers:$false

# Throttling
Write-Progress -Activity $collectionActivity -Status:"Throttling" -PercentComplete:85
Run Get-ThrottlingPolicy 
Run Get-ThrottlingPolicyAssociation 

# misc
Write-Progress -Activity $collectionActivity -Status:"Misc" -PercentComplete:85
Run Get-MigrationConfig 
Run Get-MigrationEndpoint 
Run Get-NetworkConnectionInfo -Servers:$directAccessServers -Identifier:Identity -ExecuteWhenNoServers:$false
Run Get-ProcessInfo -Servers:$directAccessServers -Identifier:TargetMachine -ExecuteWhenNoServers:$false
Run Get-OutlookProtectionRule 
Run Get-PolicyTipConfig
Run Get-RbacDiagnosticInfo 
Run Get-RoleAssignmentPolicy 
Run Get-SearchDocumentFormat 
Run Get-MailboxAuditBypassAssociation 
Run Get-SettingOverride
Run "Get-Mailbox -Arbitration" -PassThru | Save-Object -Name 'Mailbox-Arbitration'
Run "Get-Mailbox -Monitoring" -PassThru | Save-Object -Name 'Mailbox-Monitoring'
Run Get-UMService
Run "Get-SPN -Path:$Path"
Run "Get-WmiObject Win32_NetworkAdapterConfiguration" -Servers:$directAccessServers  -Identifier:ComputerName -ExecuteWhenNoServers:$false -PassThru | 
    Where {$_.IPEnabled} | Save-Object -Name:"Win32_NetworkAdapterConfiguration"

#Run "Get-WmiObject Win32_Process" -Servers:$directAccessServers -Identifier:ComputerName -ExecuteWhenNoServers:$false

# FIPS
Run Get-MalwareFilteringServer
Run Get-MalwareFilterPolicy
Run Get-MalwareFilterRule
if ($IncludeFIPS)
{
    Write-Progress -Activity $collectionActivity -Status:"FIPS" -PercentComplete:85
    Run-FIPSCommands -Servers ($directAccessServers | where {$_.IsE15OrLater -and $_.IsHubTransportServer})
}

# .NET Framework Versions
Run Get-DotNetVersion -Servers:($directAccessServers) -Identifier:Computer -ExecuteWhenNoServers:$false

# TLS Settings
Run Get-TlsRegistry -Servers $directAccessServers -Identifier:Computer -ExecuteWhenNoServers:$false

# TCPIP6 
Run Get-TCPIP6Registry -Servers $directAccessServers -Identifier:Computer -ExecuteWhenNoServers:$false

# MSInfo32
# Get-MSInfo32 -Servers $directAccessServers

# Computer Settings
# Win32_powerplan is available in Win7 & above.
Run 'Get-WmiObject -namespace "root\cimv2\power" -class Win32_PowerPlan' -Servers $directAccessServers -Identifier ComputerName -ExecuteWhenNoServers $false -PassThru | Save-Object -Name Win32_PowerPlan
Run 'Get-WmiObject -Class Win32_PageFileSetting' -Servers $directAccessServers -Identifier ComputerName -ExecuteWhenNoServers $false -PassThru | Save-Object -Name Win32_PageFileSetting
Run 'Get-WmiObject -Class Win32_ComputerSystem' -Servers $directAccessServers -Identifier ComputerName -ExecuteWhenNoServers $false -PassThru | Save-Object -Name Win32_ComputerSystem
Run 'Get-WmiObject -Class Win32_OperatingSystem' -Servers $directAccessServers -Identifier ComputerName -ExecuteWhenNoServers $false -PassThru | Save-Object -Name Win32_OperatingSystem


# Ldife for Exchange Org
Write-Progress -Activity $collectionActivity -Status:"Running Ldifde" -PercentComplete:90
Write-Log "Running Run-Ldifde -Path:$Path"
Run-Ldifde -Path:$Path | Save-Object


# Collect EventLogs
if ($IncludeEventLogs -or $IncludeEventLogsWithCrimson)
{
    Write-Progress -Activity $collectionActivity -Status:"Event Logs" -PercentComplete:90

    $eventLogPath = Join-Path $Path -ChildPath 'EventLogs'
    if ($IncludeEventLogsWithCrimson)
    {
        Collect-EventLogs -Path:$eventLogPath -Computers:$directAccessServers -IncludeCrimsonLogs
    }
    else
    {
        Collect-EventLogs -Path:$eventLogPath -Computers:$directAccessServers
    }
}

} # end of try for transcript
finally
{
    # release transcript file even when script is stopped in the middle.
    if ($transcriptEnabled)
    {        
        Stop-Transcript
    }
}

Write-Progress -Activity $collectionActivity -Status:"Packing into a Zip file" -PercentComplete:95

Write-Log "Running Zip-Folder -Path:$Path -ZipFileName:$OrgName"
Zip-Folder -Path:$Path -ZipFileName:$OrgName -RemoveFiles:(-not $KeepOutputFiles)

Write-Progress -Activity $collectionActivity -Status:"Completed" -Completed
Write-Host "Done!"
