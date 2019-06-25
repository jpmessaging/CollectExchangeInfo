# Overview
Collect-ExchangeInfo.ps1 is a PowerShell script to collect Microsoft Exchange Server related information.  This includes various Exchange-related cmdlets, event logs, and IIS configuration etc.  

For more details, see the script's comment-based help.

[Download Link](https://github.com/jpmessaging/CollectExchangeInfo/releases/download/v2019-05-31/Collect-ExchangeInfo.ps1)

# Sample usage
1. This will collect only Active Directory-based information  

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo
    ```

2. In addition to the information gathered by 1., this will include machine-specific informtion for servers whose name matches ("EX-*").  Their event logs + Exchange's crimson logs are also collected.

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo -Servers:EX-* -IncludeEventLogsWithCrimson
    ```

