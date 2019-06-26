# 概要
Collect-ExchangeInfo.ps1 は、様々な Exchange サーバーに関する構成情報を取得し、指定したパスにログ ファイルを出力します。設定変更などは実施しません。

本スクリプトは Active Directory 上の構成情報、そして Servers パラメーターで指定したサーバー上の情報を取得するものであるため、各サーバーで実施いただく必要はありません。任意のサーバー (※) で一度のみ実施ください。

※ 異なるバージョンの Exchange サーバーが混在する環境においては、最も新しいバージョンのサーバーで実行ください (例: Exchange 2013 と Exchange 2010 の混在環境の場合には、Exchange 2013 サーバーで実行ください)。

より詳細については、スクリプトのコメント ベースのヘルプを参照ください。

[ダウンロード](https://github.com/jpmessaging/CollectExchangeInfo/releases/download/v2019-05-31/Collect-ExchangeInfo.ps1)

# 実行例
1.  Active Directory 上の構成情報のみを取得します。

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo
    ```
  
2. Active Directory 上の構成情報に加えて、名前が "EX-*" にマッチするサーバーについてはそのサーバー固有の情報と、イベントログ (Exchange のクリムゾン ログ含む) を取得します。  

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo -Servers:EX-* -IncludeEventLogsWithCrimson
    ```

# 補足
ps1 ファイルをダウンロード後、以下の手順でブロックを解除します。
1. ファイルを右クリックして、プロパティを開きます
2. [全般] タブにて、「このファイルは他のコンピューターから取得したものです。このコンピューターを保護するため、このファイルへのアクセスはブロックされる可能性があります。」というメッセージが表示されている場合には、[許可する] にチェックを入れます。
