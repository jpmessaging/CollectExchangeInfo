# 概要
Collect-ExchangeInfo.ps1 は、様々な Exchange サーバーに関する構成情報を取得し、指定したパスにログ ファイルを出力します。設定変更などは実施しません。

詳細については、スクリプトのコメント ベースのヘルプを参照ください。


# 利用方法
1. Collect-ExchangeInfo.ps1 をダウンロードし、ブロックを解除します。

    [ダウンロード](https://github.com/jpmessaging/CollectExchangeInfo/releases/download/v2020-06-20/Collect-ExchangeInfo.ps1)

   1. ファイルを右クリックして、プロパティを開きます  
   2. [全般] タブにて、「このファイルは他のコンピューターから取得したものです。このコンピューターを保護するため、このファイルへのアクセスはブロックされる可能性があります。」というメッセージが表示されている場合には、[許可する] にチェックを入れます。

2. Exchange サーバー上に Collect-ExchangeInfo.ps1 をコピーします。

    本スクリプトは Active Directory 上の構成情報、そして Servers パラメーターで指定したサーバー上の情報を取得するものであるため、各サーバーで実施いただく必要はありません。任意のサーバー (※) で一度のみ実施ください。

    ※ 異なるバージョンの Exchange サーバーが混在する環境においては、最も新しいバージョンのサーバーで実行ください (例: Exchange 2013 と Exchange 2010 の混在環境の場合には、Exchange 2013 サーバーで実行ください)。

3. 管理者権限で Exchange 管理シェルを起動します。
4. スクリプトを実行します。

    採取する対象サーバーやパラメータについてはエンジニアからの案内をご確認ください。

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path <出力先フォルダー パス> -Servers:<マシン固有情報を採取する対象のサーバー>
    ```

    例: 
    ```PowerShell
    C:\Script>.\Collect-ExchangeInfo.ps1 -Path C:\Script\logs -Servers:ex01,ex02 -IncludeEventLogs 
    ```

スクリプトが終了すると、"<組織名_採取日時>.zip" が出力されます。こちらの ZIP ファイルを弊社までお寄せください。


# 実行例
1.  Active Directory 上の構成情報のみを取得します。

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo
    ```
  
2. Active Directory 上の構成情報に加えて、名前が "EX-*" にマッチするサーバーについてはそのサーバー固有の情報と、イベントログ (Exchange のクリムゾン ログ含む) を取得します。  

    ```PowerShell
    .\Collect-ExchangeInfo.ps1 -Path C:\exinfo -Servers:EX-* -IncludeEventLogsWithCrimson
    ```

# ライセンス
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。