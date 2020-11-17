---
page_type: sample
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
---
# Obtenha dados do OneDrive usando o Microsoft Graph e o MSAL.NET em um Suplemento do Office 

Aprenda a criar um suplemento do Microsoft Office que se conecte ao Microsoft Graph, encontre as três primeiras pastas de trabalho armazenadas no OneDrive for Business, busca seus nomes de arquivo e insira os nomes em um documento do Office usando o Office.js.

## Recursos
A integração de dados de provedores de serviço online aumenta o valor e a adoção de seus suplementos. O código a seguir mostra como conectar seu suplemento ao Microsoft Graph. Use este exemplo de código para:

* Conectar-se ao Microsoft Graph a partir de um Suplemento do Office.
* Use a Biblioteca MSAL.NET para implementar a estrutura de autorização do OAuth 2.0 em um suplemento.
* Use as APIs REST do OneDrive a partir do Microsoft Graph.
* Exiba uma caixa de diálogo usando o namespace da interface do usuário do Office.
* Crie um Suplemento usando ASP.NET MVC, MSAL 3.x.x para NET e Office.js. 
* Use comandos de suplemento no suplemento.

## Aplicável a

-  Excel no Windows (compra única e assinatura)
-  PowerPoint no Windows (compra única e assinatura)
-  Word no Windows (compra única e assinatura)

## Pré-requisitos

Para executar este exemplo de código, são necessários.

* Visual Studio 2019 ou posterior.

* SQL Server Express (não é mais instalado automaticamente com versões recentes do Visual Studio).

* Uma conta do Office 365 que você pode obter ingressando no [Programa para Desenvolvedores do Office 365](https://aka.ms/devprogramsignup) que inclui uma assinatura gratuita de 1 ano do Office 365.

* Pelo menos três pastas de trabalho do Excel armazenadas no OneDrive for Business na sua assinatura do Office 365.

* Office para Windows, versão 16.0.6769.2001 ou posterior.

* [Ferramentas para Desenvolvedores do Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Um Locatário do Microsoft Azure. Este suplemento requer o Azure Active Directiory (AD). O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização. Você pode adquirir uma assinatura de avaliação aqui: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solução

Solução | Autor(es)
---------|----------
suplemento do Office Microsoft Graph ASP.NET | Microsoft

## Histórico de versão

Versão | Data | Comentários
---------| -----| --------
1.0 | 8 de julho de 2019 | Versão inicial

## Aviso de isenção de responsabilidade

**ESSE CÓDIGO É FORNECIDO *NAS CIRCUNTÂNCIAS ATUAIS*SEM GARANTIA DE QUALQUER TIPO, SEJA EXPLÍCITA OU IMPLÍCITA, INCLUINDO QUAISQUER GARANTIAS IMPLÍCITAS DE ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA, COMERCIABILIDADE OU NÃO VIOLAÇÃO.**

----------

## Compile e execute a solução.

### Configurar a solução

1. No **Visual Studio**, escolha o projeto **Office-suplemento-Microsoft-Graph-ASPNETWeb**. Em **Propriedades**, certifique-se de que o **SSL Habilitado** seja **Verdadeiro**. Verifique se o **URL SSL** usa o mesmo nome de domínio e número da porta que estão listados no próximo passo.
 
2. Registre o seu aplicativo usando o [Portal de Gerenciamento do Azure](https://manage.windowsazure.com). **Faça logon com a identidade de um administrador da sua locação do Office 365 para garantir que você esteja trabalhando em um Azure Active Directory associado a essa locação.** Para aprender como registrar seus aplicativos, confira[Registrando um aplicativo na Microsoft Identity Platform](https://docs.microsoft.com/graph/auth-register-app-v2). Use as seguintes configurações:

 - REDIRECIONE O URI: https://localhost:44301/AzureADAuth/Authorize
 - TIPOS DE CONTA COM SUPORTE: “Apenas contas neste diretório organizacional”
 - CONCESSÃO IMPLÍCITA: Não ative nenhuma opção de Concessão Implícita
 - PERMISSÕES de API (permissões delegadas, não permissões de aplicativo): **Files.Read.All** e **User.Read**

	> Observação: Após registrar o seu aplicativo, copie a **ID do Aplicativo (cliente)** e a **ID do Diretório (locatário)** na folha **Visão geral** do Registro de Aplicativo no Portal de Gerenciamento do Azure. Ao criar o segredo do cliente na folha **Certificados e segredos**, copie-o também. 
	 
3.  No web.config, use os valores que você copiou na etapa anterior. Defina **AAD: ClientID** para a ID do cliente, defina **AAD: ClientSecret** para o seu segredo de cliente e defina **"AAD: O365TenantID"** à sua ID de locatário. 

### Executar a solução

1. Abra o arquivo de solução do Visual Studio. 
2. Clique com o botão direito do mouse **Office-suplemento-Microsoft-Graph-ASPNET** solução no **Gerenciador de Soluções** (não os nós do projeto), em seguida, escolha **Configurar projetos de inicialização**. Marque a caixa de seleção **vários projetos de inicialização**. Verifique se o projeto que termina com "Web" está listado primeiro.
3. No menu **Compilar**, selecione **Solução Limpa**. Quando terminar, abra o menu **Compilar** novamente e selecione **Compilar Solução**.
4. No **Gerenciador de soluções**, selecione o nó do projeto **Suplemento-Office-Microsoft-Graph-ASPNET** (não o primeiro nó da solução e não o projeto cujo nome termina em "Web").
5. No painel **Propriedades**, abra o menu suspenso Iniciar Documento e escolha uma das três opções (Excel, Word ou PowerPoint).

    ![Escolha o aplicativo host do Office desejado: Excel ou PowerPoint ou Word](images/SelectHost.JPG)

6. Pressione F5. 
7. No aplicativo do Office, escolha **Inserir** > **Abrir Suplemento** no grupo **Arquivos do OneDrive** para abrir o suplemento do painel de tarefas.
8. As páginas e os botões do suplemento são auto-explicativos. 

## Problemas conhecidos

* O controle giratório do Fabric só aparece brevemente ou nem isso.

## Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode enviar seus comentários na seção *Problemas* deste repositório.
Perguntas sobre o desenvolvimento de suplementos do Office devem ser publicadas em [Stack Overflow](http://stackoverflow.com). Certifique-se de que as suas perguntas estejam marcadas com [office-js], [MicrosoftGraph].

## Recursos adicionais

* [Documentação do Microsoft Graph](https://docs.microsoft.com/graph/)
* [Documentação de Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Direitos autorais
Direitos autorais (c) 2019 Microsoft Corporation. Todos os direitos reservados.

Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
