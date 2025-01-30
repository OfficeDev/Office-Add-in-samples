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
# 使用 Microsoft Graph 和 MSAL.NET 在 Office 外接程序中获取 OneDrive 数据 

了解如何构建连接到 Microsoft Graph 的 Microsoft Office 外接程序，查找存储在 OneDrive for Business 中的前三个工作簿，提取其文件名，然后使用 Office.js 将名称插入 Office 文档。

## 功能
集成来自联机服务提供程序的数据可提高外接程序的价值和采用率。此代码示例演示了如何将外接程序连接到 Microsoft Graph。使用此代码示例可执行以下操作：

* 从 Office 外接程序连接到 Microsoft Graph。
* 使用 MSAL.NET 库在外接程序中实现 OAuth 2.0 授权框架。
* 从 Microsoft Graph 中使用 OneDrive REST API。
* 使用 Office UI 命名空间显示对话框。
* 使用 ASP.NET MVC、适用于 .NET 的 MSAL 3.x.x 和 Office.js 构建外接程序。 
* 在外接程序中使用外接程序命令。

## 适用于

-  Windows 版 Excel（一次性购买和订阅）
-  Windows 版 PowerPoint（一次性购买和订阅）
-  Windows 版 Word（一次性购买和订阅）

## 先决条件

必须符合以下条件才能运行此代码示例。

* Visual Studio 2019 或更高版本。

* SQL Server Express（不再随最新版本的 Visual Studio 一起自动安装。）

* Office 365 帐户，获取方法为加入 [Office 365 开发人员计划](https://aka.ms/devprogramsignup)，其中包含为期 1 年的免费 Office 365 订阅。

* 在 Office 365 订阅的 OneDrive for Business 中存储的至少三个 Excel 工作簿。

* Windows 版 Office，版本 16.0.6769.2001 或更高版本。

* [Office 开发人员工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* 一个 Microsoft Azure 租户。此外接程序需要 Azure Active Directiory (AD)。Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。你还可在此处获得试用订阅：[Microsoft Azure](https://account.windowsazure.com/SignUp)。

## 解决方案

解决方案 | 作者
---------|----------
Office 外接程序 Microsoft Graph ASP.NET | Microsoft

## 版本历史记录

版本 | 日期 | 批注
---------| -----| --------
1.0 | 2019 年 7 月 8 日| 初始发行版

## 免责声明

**此代码*按原样提供*，不提供任何明示或暗示的担保，包括对特定用途适用性、适销性或不侵权的默示担保。**

----------

## 构建和运行解决方案

### 配置解决方案

1. 在 **Visual Studio** 中，选择**“Office-Add-in-Microsoft-Graph-ASPNETWeb”**项目。在**“属性”**中，确保**“已启用 SSL”**为**“True”**。验证 **SSL URL** 使用的域名和端口号与下一步中列出的相同。
 
2. 使用 [Azure 管理门户](https://manage.windowsazure.com)注册你的应用程序。**使用 Office 365 租赁的管理员标识登录，以确保正在使用与该租赁相关联的 Azure Active Directory。**若要了解如何注册应用程序，请参阅 [向 Microsoft 标识平台注册应用程序](https://learn.microsoft.com/graph/auth-register-app-v2)。使用以下设置：

 - 重定向 URI：https://localhost:44301/AzureADAuth/Authorize
 - 支持的帐户类型：“仅限此组织目录中的帐户”
 - 隐式授权：不启用任何隐式授权选项
 - API 权限（代理权限，而不是应用程序权限）：**Files.Read.All** 和 **User.Read**

	> 注意：注册应用程序之后，复制 Azure 管理门户的**“概览”**部分上的**“应用程序(客户端) ID”**和**“目录(租户) ID”**。在**“证书和密码”**部分创建客户端密码时，同样复制该密码。 
	 
3.  在 web.config 中，使用你在上一步中复制的值。将**“AAD:ClientID”**设置为客户端 ID，将**“AAD:ClientSecret”**设置为客户端密码，并将**“AAD:O365TenantID”**设置为租户 ID。 

### 运行解决方案

1. 打开 Visual Studio 解决方案文件。 
2. 在**解决方案资源管理器**（而不是项目节点）中右键单击 **Office-Add-in-Microsoft-Graph-ASPNET** 解决方案，然后选择“**设置启动项目**”。选择“**多启动项目**”单选按钮。请确保先列出以“Web”结尾的项目。
3. 在“**生成**”菜单上，选择“**清理解决方案**”。完成后，再次打开“**生成**”菜单，并选择“**生成解决方案**”。
4. 在“**解决方案资源管理器**”中，选择“**Office-Add-in-Microsoft-Graph-ASPNET**”项目节点（而不是顶部的解决方案节点，也不是名称以“Web”结尾的项目）。
5. 在“**属性**”窗格中，打开“**启动文档**”下拉列表，然后选择三个选项之一（“Excel”、“Word”或“PowerPoint”）。

    ![选择所需的 Office 主机应用程序：](images/SelectHost.JPG)Excel、PowerPoint 或 Word](images/SelectHost.JPG)

6. 按 <kbd>F5</kbd>。 
7. 在 Office 应用程序中，选择“**OneDrive文件**”组中的“**插入**”>“**打开外接程序**”，打开任务窗格外接程序。
8. 外接程序中的页面和按钮一目了然。 

## 已知问题

* 结构微调控件仅暂时显示或根本不显示。

## 问题和意见

我们乐意倾听你对此示例的反馈。可以在此存储库中的*“问题”*部分向我们发送反馈。
与开发 Office 外接程序相关的问题应发布到[堆栈溢出](http://stackoverflow.com)。确保你的问题使用了 [office-js] 和 [MicrosoftGraph] 标记。

## 其他资源

* [Microsoft Graph 文档](https://learn.microsoft.com/graph/)
* [Office 外接程序文档](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## 版权信息
版权所有 (c) 2019 Microsoft Corporation。保留所有权利。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
