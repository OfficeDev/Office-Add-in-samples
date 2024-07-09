---
page_type: sample
products:
  - office
  - office-excel
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
---
# Office アドインの Microsoft Graph と MSAL.NET を使用して OneDrive データを取得する 

Microsoft Graph に接続して OneDrive for Business に保存されている最初の 3 つのブックを検索し、それらのファイル名を取得して、Office.js. を使用してその名前を Office ドキュメントに挿入できるよう Microsoft Outlook アドインの作成方法について学習します。

## 機能
オンライン サービス プロバイダーからのデータを統合すると、アドインの価値が向上し、採用できる機会が増えます。このコード サンプルでは、Microsoft Graph にアドインを接続する方法を示します。このコード サンプルを使用して、以下を実行します。

* Office アドインから Microsoft Graph に接続します。
* MSAL.NET ライブラリを使用して、アドインに OAuth 2.0 承認フレームワークを実装します。
* Microsoft Graph から OneDrive REST API を使用します。
* Office UI 名前空間を使用してダイアログを表示します。
* ASP.NET MVC、MSAL 3.x.x for .NET、Office.js を使用してアドインをビルドします。 
* アドインでアドイン コマンドを使用します。

## 適用対象

-  Windows 上の Excel (1 回限りの購入とサブスクリプション)
-  Windows 上の PowerPoint (1 回限りの購入とサブスクリプション)
-  Windows 上の Word (1 回限りの購入とサブスクリプション)

## 前提条件

このコード サンプルを実行するには、以下が必要です。

* Visual Studio 2019 以降。

* SQL Server Express (最新バージョンの Visual Studio では自動的にインストールされなくなりました。)

* [Office 365 開発者プログラム](https://aka.ms/devprogramsignup)に参加すると取得できる Office 365 アカウント。Office 365 の 1 年間の無料サブスクリプションが含まれています。

* Office 365 サブスクリプションの OneDrive for Business に保存された少なくとも 3 つの Excel ワークブック。

* Windows 上の Office のバージョン 16.0.6769.2001 以降。

* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure テナント。このアドインには、Azure Active Directory (AD) が必要です。Azure AD は、アプリケーションでの認証と承認に使う ID サービスを提供します。ここでは、試用版サブスクリプションを取得できます。[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## ソリューション

ソリューション | 作成者
---------|----------
Office アドイン Microsoft Graph ASP.NET | Microsoft

## バージョン履歴

バージョン | 日付 | コメント
---------| -----| --------
1.0 | 2019 年 7 月 8 日 | 初期リリース

## 免責事項

**このコードは、明示または黙示のいかなる種類の保証なしに*現状のまま*提供されるものであり、特定目的への適合性、商品性、権利侵害の不存在についての暗黙的な保証は一切ありません。**

----------

## ソリューションの構築と実行

### ソリューションを構成する

1. **Visual Studio** で、**Office-Add-in-Microsoft-Graph-ASPNETWeb** プロジェクトを選択します。**[プロパティ]** で、**[SSL が有効]** が **True** であることを確認します。**[SSL URL]** で、次の手順でリストされているのと同じドメイン名とポート番号が使用されていることを確認します。
 
2. [Azure の管理ポータル](https://manage.windowsazure.com)を使用してアプリケーションを登録します。**Office 365 テナントの管理者の ID でログインして、そのテナントに関連付けられている Azure Active Directory で作業していることを確認します。**アプリケーションの登録の方法については、「[Microsoft ID プラットフォームにアプリケーションを登録する](https://learn.microsoft.com/graph/auth-register-app-v2)」を参照してください。次に示す設定を使用します。

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize
 - サポートされているアカウントの種類:"この組織のディレクトリ内のアカウントのみ"
 - 暗黙的な付与:暗黙的な付与オプションを有効にしない
 - API アクセス許可 (委任されたアクセス許可、アプリケーション アクセス許可ではありません):**Files.Read.All** と **User.Read**

	> 注:注: アプリケーションを登録したら、Azure の管理ポータルにある [アプリの登録] の **[概要]** ブレードの**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** をコピーします。**[証明書とシークレット]** ブレードでクライアント シークレットを作成したら、それもコピーします。 
	 
3.  web.config で、前の手順でコピーした値を使用します。**[AAD:ClientID]** にクライアント ID、**[AAD:ClientSecret]** にクライアント シークレット、**[AAD:O365TenantID]** にテナント ID を設定します。 

### ソリューションを実行する

1. Visual Studio ソリューション ファイルを開きます。 
2. [**ソリューション エクスプローラー**] (プロジェクト ノードではありません) で、[**Office-Add-in-Microsoft-Graph-ASPNET**] ソリューションを右クリックし、**[スタートアップ プロジェクトの設定]** を選択します。[**マルチ スタートアップ プロジェクト**] ラジオ ボタンを選択します。最後に「Web」で終わるプロジェクトが表示されていることを確認します。
3. [**ビルド**] メニューで [**ソリューションのクリーン**] を選択します。終了したら、[**ビルド**] メニューをもう一度開き、[**ソリューションのビルド**] を選択します。
4. [**ソリューション エクスプローラー**] で、[**Office-Add-in-Microsoft-Graph-ASPNET**] を選択します (一番上のソリューション ノードではなく、「Web」で終わる名前のプロジェクトではありません)。
5. [**プロパティ**] ウィンドウで、[**ドキュメントの開始**] ドロップダウンを開き、3 つのオプション (Excel、Word、または PowerPoint) のいずれかを選択します。

    ![必要な Office ホスト アプリケーションを選択する:Excel、PowerPoint、または Word](images/SelectHost.JPG)

6. F5 キーを押します。 
7. Office アプリケーションで、[**OneDrive ファイル**] グループから [**挿入**]、[**アドインを開く**] の順に選択して、タスク ウィンドウのアドインを開きます。
8. アドインのページとボタンは、わかりやすく説明不要です。 

## 既知の問題

* ファブリック スピナー制御が、わずかに表示されるか、まったく表示されません。

## 質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。
Office アドインの開発に関する質問は、「[Stack Overflow](http://stackoverflow.com)」に投稿してください。質問には、[office-js]、および [MicrosoftGraph] のタグを付けてください。

## その他のリソース

* [Microsoft Graph ドキュメント](https://learn.microsoft.com/graph/)
* [Office アドイン ドキュメント](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## 著作権
Copyright (c) 2019 Microsoft Corporation.All rights reserved.

このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
