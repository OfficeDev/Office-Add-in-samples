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
description: "了解如何构建连接到 Microsoft Graph 的 Microsoft Outlook 外接程序。"
---

# 使用 Microsoft Graph 和 MSAL 在 Outlook 外接程序中获取 Excel 工作簿 

了解如何构建连接到 Microsoft Graph 的 Microsoft Outlook 外接程序，查找存储在 OneDrive for Business 中的前三个工作簿，提取其文件名，然后将名称插入 Outlook 中的新邮件撰写表单。

## 功能

集成来自联机服务提供程序的数据可提高外接程序的价值和采用率。此代码示例演示了如何将 Outlook 外接程序连接到 Microsoft Graph。使用此代码示例可执行以下操作：

* 从 Office 外接程序连接到 Microsoft Graph。
* 使用 MSAL .NET 库在外接程序中实现 OAuth 2.0 授权框架。
* 从 Microsoft Graph 中使用 OneDrive REST API。
* 使用 Office UI 命名空间显示对话框。
* 使用 ASP.NET MVC、适用于 .NET 的 MSAL 3.x.x 和 Office.js 构建外接程序。 

## 适用于

-  所有平台上的 Outlook

## 先决条件

必须符合以下条件才能运行此代码示例。

* Visual Studio 2019 或更高版本。

* SQL Server Express（如果不随最新版本的 Visual Studio 一起自动安装。）

* Office 365 帐户，获取方法为加入 [Office 365 开发人员计划](https://aka.ms/devprogramsignup)，其中包含为期 1 年的免费 Office 365 订阅。

* 在 Office 365 订阅的 OneDrive for Business 中存储的至少三个 Excel 工作簿。

* 如果要在桌面而不是 Outlook Online 上进行调试，则是可选的：Windows 版 Outlook，版本 1809 或更高版本。
* [Office 开发人员工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* 一个 Microsoft Azure 租户。此外接程序需要 Azure Active Directiory (AD)。Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。你还可在此处获得试用订阅：[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## 解决方案

解决方案 | 作者
---------|----------
Outlook 外接程序 Microsoft Graph ASP.NET | Microsoft

## 版本历史记录

版本 | 日期 | 批注
---------| -----| --------
1.0 | 2019 年 7 月 8 日| 初始发行版

## 免责声明

**此代码*按原样提供*，不提供任何明示或暗示的担保，包括对特定用途适用性、适销性或不侵权的默示担保。**

----------

## 构建和运行解决方案

## 配置解决方案

1. 在 **Visual Studio** 中，选择**“Outlook-Add-in-Microsoft-Graph-ASPNETWeb”**项目。在**“属性”**中，确保**“已启用 SSL”**为**“True”**。验证 **SSL URL** 使用的域名和端口号与下一步中列出的相同。
 
2. 使用 [Azure 管理门户](https://manage.windowsazure.com)注册你的应用程序。**使用 Office 365 租赁的管理员标识登录，以确保正在使用与该租赁相关联的 Azure Active Directory。**若要了解如何注册应用程序，请参阅 [向 Microsoft 标识平台注册应用程序](https://docs.microsoft.com/graph/auth-register-app-v2)。使用以下设置：

 - 重定向 URI：https://localhost:44301/AzureADAuth/Authorize
 - 支持的帐户类型：“仅限此组织目录中的帐户”
 - 隐式授权：不启用任何隐式授权选项
 - API 权限（代理权限，而不是应用程序权限）：**Files.Read.All** 和 **User.Read**

	> 注意：注册应用程序之后，复制 Azure 管理门户的**“概览”**部分上的**“应用程序(客户端) ID”**和**“目录(租户) ID”**。在**“证书和密码”**部分创建客户端密码时，同样复制该密码。 
	 
3.  在 web.config 中，使用你在上一步中复制的值。将**“AAD:ClientID”**设置为客户端 ID，将**“AAD:ClientSecret”**设置为客户端密码，并将**“AAD:O365TenantID”**设置为租户 ID。 

## 运行解决方案

1. 打开 Visual Studio 解决方案文件。 
2. 在**解决方案资源管理器**（而不是项目节点）中右键单击 **Outlook-Add-in-Microsoft-Graph-ASPNET** 解决方案，然后选择“**设置启动项目**”。选择“**多启动项目**”单选按钮。请确保先列出以“Web”结尾的项目。
3. 在“**生成**”菜单上，选择“**清理解决方案**”。完成后，再次打开“**生成**”菜单，并选择“**生成解决方案**”。
4. 在“**解决方案资源管理器**”中，选择“**Outlook-Add-in-Microsoft-Graph-ASPNET**”项目节点（而不是顶部的解决方案节点，也不是名称以“Web”结尾的项目）。
5. 在**“属性”**窗格中，打开**“启动操作”**下拉列表，然后选择在桌面 Outlook 中，还是在列出的浏览器之一中的 Outlook 网页版中运行外接程序。（*请勿选择“Internet Explorer”。有关原因，请参阅以下**已知问题**。*） 

    ![选择所需的 Oulook 主机：台式机或浏览器之一](images/StartAction.JPG)

6. 按 F5。首次执行此操作时，系统将提示你指定用于调试外接程序的用户的电子邮件和密码。使用你的 O365 租户的管理员凭据。 

    ![带用户电子邮件和密码文本框的窗体](images/CredentialsPrompt.JPG)

    >注意：浏览器将打开 Office 网页版的登录页面。（因此，如果这是首次运行该外接程序，则将输入两次用户名和密码。） 

剩余步骤取决于你是在桌面 Outlook 还是 Outlook 网页版中运行外接程序。

### 在 Outlook 网页版中运行解决方案

1. 将在浏览器窗口中打开 Outlook 网页版。在 Outlook​​ 中，单击**“新建”**新建电子邮件。 
2. 撰写窗体下面是一个工具栏，包含用于**“发送”**、**“放弃”**和其他实用工具的按钮。根据你正在使用的 **Outlook 网页版**体验，该外接程序的图标位于此工具栏最右端附近，或者位于你单击此工具栏上的 **...** 按钮时将打开的下拉菜单上。

   ![“插入文件外接程序”的图标](images/Onedrive_Charts_icon_16x16px.png)

3. 单击此图标，打开任务窗格外接程序。
4. 使用外接程序将用户的 OneDrive 帐户中的前三个工作簿的名称添加到邮件中。外接程序的页面和按钮一目了然。

## 在桌面版 Outlook 中运行项目。

1. 桌面版 Outlook 将打开。在 Outlook​​ 中，单击**“新建电子邮件”**新建电子邮件。 
2. 在**“邮件”**窗体的**“邮件”**功能区上，在名为“**OneDrive 文件**”的组中有一个标为“**打开外接程序**”的按钮。单击该按钮，打开外接程序。
3. 使用外接程序将用户的 OneDrive 帐户中的前三个工作簿的名称添加到邮件中。外接程序的页面和按钮一目了然。

## 已知问题

* 结构微调控件仅暂时显示或根本不显示。 
* 如果你在 Internet Explorer 中运行，则当你尝试登录时，将收到一条错误消息，提示你必须将 `https://localhost:44301` 和 `https://outlook.office.com`（或者 `https://outlook.office365.com`）放在相同的安全区域中。但是即使你这样做，也会发生此错误。 

## 问题和意见

我们希望得到你对*“使用 Microsoft Graph 和 MSAL 在 Outlook 外接程序中获取 Excel 工作簿”* 示例的相关反馈。可以在此存储库中的*“问题”*部分向我们发送反馈。
与 Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)。确保你的问题使用了 [office-js]、[MicrosoftGraph] 和 [API] 标记。

## 其他资源

* [Microsoft Graph 文档](https://docs.microsoft.com/graph/)
* [Office 外接程序文档](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## 版权信息
版权所有 (c) 2019 Microsoft Corporation。保留所有权利。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
