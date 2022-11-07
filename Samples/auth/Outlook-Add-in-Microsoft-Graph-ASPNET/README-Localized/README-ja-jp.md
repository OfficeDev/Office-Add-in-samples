---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
description: "Microsoft Graph に接続する Microsoft Outlook アドインを構築する方法を学習します。"
---

# Outlook アドインで Microsoft Graph と MSAL を使用して Excel ブックを取得する 

Microsoft Graph に接続し、OneDrive for Business に保存されている最初の 3 つのブックを検索して、それらのファイル名を取得して Outlook の新しいメッセージの作成フォームにその名前を挿入できるよう Microsoft Outlook アドインの作成方法について学習します。

## 機能

オンライン サービス プロバイダーからのデータを統合すると、アドインの価値が向上し、採用できる機会が増えます。このコード サンプルでは、Microsoft Graph に Outlook アドインを接続する方法を示します。このコード サンプルを使用して、以下を実行します。

* Office アドインから Microsoft Graph に接続します。
* MSAL .NET ライブラリを使用して、アドインに OAuth 2.0 承認フレームワークを実装します。
* Microsoft Graph から OneDrive REST API を使用します。
* Office UI 名前空間を使用してダイアログを表示します。
* ASP.NET MVC、MSAL 3.x.x for .NET、Office.js を使用してアドインをビルドします。 

## 適用対象

-  すべてのプラットフォームの Outlook

## 前提条件

このコード サンプルを実行するには、以下が必要です。

* Visual Studio 2019 以降。

* SQL Server Express (最新バージョンの Visual Studio で自動的にインストールされない場合。)

* [Office 365 開発者プログラム](https://aka.ms/devprogramsignup)に参加すると取得できる Office 365 アカウント。Office 365 の 1 年間の無料サブスクリプションが含まれています。

* Office 365 サブスクリプションの OneDrive for Business に保存された少なくとも 3 つの Excel ワークブック。

* Outlook Online の代わりにデスクトップでデバッグする場合は、省略可能です。Outlook for Windows のバージョン 1809 以降。
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Microsoft Azure テナント。このアドインには、Azure Active Directory (AD) が必要です。Azure AD は、アプリケーションでの認証と承認に使う ID サービスを提供します。ここでは、試用版サブスクリプションを取得できます。[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## ソリューション

ソリューション | 作成者
---------|----------
Outlook Add-in Microsoft Graph ASP.NET | Microsoft

## バージョン履歴

バージョン | 日付 | コメント
---------| -----| --------
1.0 | 2019 年 7 月 8 日 | 初期リリース

## 免責事項

**このコードは、明示または黙示のいかなる種類の保証なしに*現状のまま*提供されるものであり、特定目的への適合性、商品性、権利侵害の不存在についての暗黙的な保証は一切ありません。**

----------

## ソリューションの構築と実行

## ソリューションを構成する

1. **Visual Studio** で、**Outlook-Add-in-Microsoft-Graph-ASPNETWeb** プロジェクトを選択します。**[プロパティ]** で、**[SSL が有効]** が **True** であることを確認します。**[SSL URL]** で、次の手順でリストされているのと同じドメイン名とポート番号が使用されていることを確認します。
 
2. [Azure の管理ポータル](https://manage.windowsazure.com)を使用してアプリケーションを登録します。**Office 365 テナントの管理者の ID でログインして、そのテナントに関連付けられている Azure Active Directory で作業していることを確認します。**アプリケーションの登録の方法については、「[Microsoft ID プラットフォームにアプリケーションを登録する](https://learn.microsoft.com/graph/auth-register-app-v2)」を参照してください。次に示す設定を使用します。

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize
 - サポートされているアカウントの種類:"この組織のディレクトリ内のアカウントのみ"
 - 暗黙的な付与:暗黙的な付与オプションを有効にしない
 - API アクセス許可 (委任されたアクセス許可、アプリケーション アクセス許可ではありません):**Files.Read.All** と **User.Read**

	> 注:注: アプリケーションを登録したら、Azure の管理ポータルにある [アプリの登録] の **[概要]** ブレードの**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** をコピーします。**[証明書とシークレット]** ブレードでクライアント シークレットを作成したら、それもコピーします。 
	 
3.  web.config で、前の手順でコピーした値を使用します。**[AAD:ClientID]** にクライアント ID、**[AAD:ClientSecret]** にクライアント シークレット、**[AAD:O365TenantID]** にテナント ID を設定します。 

## ソリューションを実行する

1. Visual Studio ソリューション ファイルを開きます。 
2. [**ソリューション エクスプローラー**] (プロジェクト ノードではありません) で、[**Outlook-Add-in-Microsoft-Graph-ASPNET**] ソリューションを右クリックし、**[スタートアップ プロジェクトの設定]** を選択します。[**マルチ スタートアップ プロジェクト**] ラジオ ボタンを選択します。最後に「Web」で終わるプロジェクトが表示されていることを確認します。
3. [**ビルド**] メニューで [**ソリューションのクリーン**] を選択します。終了したら、[**ビルド**] メニューをもう一度開き、[**ソリューションのビルド**] を選択します。
4. [**ソリューション エクスプローラー**] で、[**Outlook-Add-in-Microsoft-Graph-ASPNET**] を選択します (一番上のソリューション ノードではなく、「Web」で終わる名前のプロジェクトではありません)。
5. [**プロパティ**] ウィンドウで、[**操作の開始**] ドロップ ダウンを開き、表示されたブラウザーのいずれかで、デスクトップ Outlook または Outlook on the web でアドインを実行するかどうかを選択します。(*Internet Explorer は選択しないでください。理由については、以下の**既知の問題**を参照してください。*) 

    ![希望の Oulook ホスト: デスクトップまたはブラウザーのいずれかを選択します](images/StartAction.JPG)

6. F5 キーを押します。初めて実行するときに、アドインのデバッグに使用するユーザーのメールアドレスとパスワードを指定するように求められます。O365 テナントの管理者の資格情報を使用します。 

    ![ユーザーのメールアドレスとパスワードを入力するテキスト ボックスを含むフォーム](images/CredentialsPrompt.JPG)

    >注:ブラウザーが開き、Office on the web 用の [ログイン] ページが表示されます。(つまり、アドインを初めて実行する場合は、ユーザー名とパスワードを 2 回入力します)。 

残りの手順は、デスクトップ版の Outlook または Outlook on the web でアドインを実行しているかどうかによって異なります。

### Outlook on the web を使用してソリューションを実行する

1. Outlook for Web はブラウザー ウィンドウで開きます。Outlook で、[**新規**] をクリックして、新しいメール メッセージを作成します。 
2. [作成] フォームの下には、[**送信**]、[**廃棄**]、その他のユーティリティ用のボタンを含むツール バーがあります。使用している **Outlook on the web** エクスペリエンスによっては、アドイン用アイコンはこのツールバーの右端にあるか、このツール バーの [**...**] ボタンをクリックするとき開くドロップ ダウン メニューにあります。

   ![ファイル挿入アドイン用アイコン](images/Onedrive_Charts_icon_16x16px.png)

3. アイコンをクリックして、タスク ウィンドウ アドインを開きます。
4. アドインを使用して、ユーザーの OneDrive アカウントにある最初の 3 つのワークブックの名前をメッセージに追加します。アドインのページとボタンは、わかりやすく説明不要です。

## デスクトップ版の Outlook でプロジェクトを実行する

1. デスクトップ版の Outlook が開きます。Outlook で、[**新規メール**] をクリックして、新しいメール メッセージを作成します。 
2. [**メッセージ**] フォームの [**メッセージ**] リボンには、[**OneDrive ファイル**] と呼ばれるグループに [** 開いているアドイン**] とラベル付けされたボタンがあります。そのボタンをクリックして、アドインを開きます。
3. アドインを使用して、ユーザーの OneDrive アカウントにある最初の 3 つのワークブックの名前をメッセージに追加します。アドインのページとボタンは、わかりやすく説明不要です。

## 既知の問題

* ファブリック スピナー制御が、わずかに表示されるか、まったく表示されません。 
* Internet Explorer で実行している場合は、ログインしようとすると、「同じセキュリティ ゾーンに `https://localhost:44301` および `https://outlook.office.com` (または `https://outlook.office365.com`) を配置する必要があります」というエラーが表示されます。ただし、それを行っている場合でも、このエラーが発生します。 

## 質問とコメント

*Office アドインで Microsoft Graph および MSAL を使用して Excel ワークブックを取得する* サンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。
Office 365 開発全般の質問につきましては、「[Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問には、[office-js]、[MicrosoftGraph]、[API] のタグを付けてください。

## その他のリソース

* [Microsoft Graph ドキュメント](https://learn.microsoft.com/graph/)
* [Office アドイン ドキュメント](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## 著作権
Copyright (c) 2019 Microsoft Corporation.All rights reserved.

このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
