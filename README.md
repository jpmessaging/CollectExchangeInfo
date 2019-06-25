# Overview
Collect-ExchangeInfo.ps1 is a PowerShell script to collect Microsoft Exchange Server related information.  This includes various Exchange-related cmdlets, event logs, and IIS configuration etc.  

For more details, see the script's comment-based help.

[Download Link](https://github.com/jpmessaging/CollectExchangeInfo/releases/download/v2019-05-31/Collect-ExchangeInfo.ps1)

Load-Clixml.ps1 contains a function "Load-Clixml" which simply imports all xml files in a specified folder and creates corresponding variables in the global scope. For example, for a file "ExchangeServer.xml" it creates a variable $ExchangeServer.  This script is just to make it easy to load all the xml files collected.  You do not need this file to use Collect-ExchangeInfo.ps1 and this function can be used independently.

# Sample usage
1. This will collect only Active Directory-based information  

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo
    ```

2. In addition to the information gathered by 1., this will include machine-specific informtion for servers whose name matches ("EX-*").  Their event logs + Exchange's crimson logs are also collected.

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo -Servers:EX-* -IncludeEventLogsWithCrimson
    ```

# Note
After you download the ps1 file, make sure to "Unblock":

1. Right-click the ps1 file and click [Property]
2. In the [General] tab, if you see "This file came from another computer and might be blocked to help protect this computer], check [Unblock]
