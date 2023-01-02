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
description: "Saiba como criar um suplemento do Microsoft Outlook que se conecta ao Microsoft Graph"
---

# Obter pastas de trabalho do Excel usando o Microsoft Graph e MSAL em um Suplemento do Outlook 

Aprenda a criar um suplemento do Microsoft Outlook que se conecta ao Microsoft Graph, encontra as três primeiras pastas de trabalho armazenadas no OneDrive for Business, busca seus nomes de arquivo e insere os nomes em um novo formulário de redação de mensagens no Outlook.

## Recursos

A integração de dados de provedores de serviço online aumenta o valor e a adoção de seus suplementos. O código a seguir mostra como conectar seu suplemento ao Microsoft Graph. Use este exemplo de código para:

* Conectar-se ao Microsoft Graph a partir de um Suplemento do Office.
* Use a Biblioteca do MSAL .NET para implementar a estrutura de autorização do OAuth 2.0 em um suplemento.
* Use as APIs REST do OneDrive a partir do Microsoft Graph.
* Exiba uma caixa de diálogo usando o namespace da interface do usuário do Office.
* Crie um Suplemento usando ASP.NET MVC, MSAL 3.x.x para NET e Office.js. 

## Aplicável ao

-  Outlook em todas as plataformas

## Pré-requisitos

Para executar este exemplo de código, são necessários.

* Visual Studio 2019 ou posterior.

* SQL Server Express (se não for instalado automaticamente com versões recentes do Visual Studio.)

* Uma conta do Office 365 que você pode obter ingressando no [Programa para Desenvolvedores do Office 365](https://aka.ms/devprogramsignup) que inclui uma assinatura gratuita de 1 ano do Office 365.

* Pelo menos três pastas de trabalho do Excel armazenadas no OneDrive for Business na sua assinatura do Office 365.

* Opcional, se você quiser depurar na área de trabalho em vez do Outlook Online: Outlook para Windows, versão 1809 ou posterior.
* [Ferramentas para Desenvolvedores do Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Um Locatário do Microsoft Azure. Este suplemento requer o Azure Active Directiory (AD). O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização. Você pode adquirir uma assinatura de avaliação aqui: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solução

Solução | Autor(es)
---------|----------
Suplemento do Office Microsoft Graph ASP.NET | Microsoft

## Histórico de versão

Versão | Data | Comentários
---------| -----| --------
1.0 | 8 de julho de 2019 | Versão inicial

## Aviso de isenção de responsabilidade

**ESSE CÓDIGO É FORNECIDO *NAS CIRCUNTÂNCIAS ATUAIS*SEM GARANTIA DE QUALQUER TIPO, SEJA EXPLÍCITA OU IMPLÍCITA, INCLUINDO QUAISQUER GARANTIAS IMPLÍCITAS DE ADEQUAÇÃO A UMA FINALIDADE ESPECÍFICA, COMERCIABILIDADE OU NÃO VIOLAÇÃO.**

----------

## Compile e execute a solução.

## Configurar a solução

1. No **Visual Studio**, escolha o projeto **Outlook-Add-in-Microsoft-Graph-ASPNETWeb**. Em **Propriedades**, certifique-se que o **SSL Habilitado** está definido como True. Verifique se o **URL SSL** usa o mesmo nome de domínio e número da porta que estão listados no próximo passo.
 
2. Registre o seu aplicativo usando o [Portal de Gerenciamento do Azure](https://manage.windowsazure.com). **Faça logon com a identidade de um administrador da sua locação do Office 365 para garantir que você esteja trabalhando em um Azure Active Directory associado a essa locação.** Para aprender como registrar seus aplicativos, confira[Registrando um aplicativo na Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2). Use as seguintes configurações:

 - REDIRECIONE O URI: https://localhost:44301/AzureADAuth/Authorize
 - TIPOS DE CONTA COM SUPORTE: “Apenas contas neste diretório organizacional”
 - CONCESSÃO IMPLÍCITA: Não ative nenhuma opção de Concessão Implícita
 - PERMISSÕES de API (permissões delegadas, não permissões de aplicativo): **Files.Read.All** e **User.Read**

	> Observação: Após registrar o seu aplicativo, copie a **ID do Aplicativo (cliente)** e a **ID do Diretório (locatário)** na folha **Visão geral** do Registro de Aplicativo no Portal de Gerenciamento do Azure. Ao criar o segredo do cliente na folha **Certificados e segredos**, copie-o também. 
	 
3.  No web.config, use os valores que você copiou na etapa anterior. Defina **AAD: ClientID** para a ID do cliente, defina **AAD: ClientSecret** para o seu segredo de cliente e defina **"AAD: O365TenantID"** à sua ID de locatário. 

## Executar a solução

1. Abra o arquivo de solução do Visual Studio. 
2. Clique com o botão direito do mouse em **Office-suplemento-Microsoft-Graph-ASPNET** na solução **Gerenciador de Soluções** (não os nós do projeto), em seguida, escolha **Configurar projetos de inicialização**. Marque o botão de opção **Vários projetos de inicialização**. Verifique se o projeto que termina com "Web" está listado primeiro.
3. No menu **Compilar**, selecione **Solução Limpa**. Quando terminar, abra o menu **Compilar** novamente e selecione **Compilar Solução**.
4. Em **Gerenciador de Soluções**, selecione o nó do projeto **Suplemento-Office-Microsoft-Graph-ASPNET** (não o primeiro nó da solução e não o projeto cujo nome termina em "Web").
5. No painel **Propriedades**, abra o menu suspenso **Iniciar Ação** e escolha se deseja executá-lo no Outlook ou no Outlook na Web em um dos navegadores listados. (*Não escolha o Internet Explorer. Confira **Problemas Conhecidos** abaixo para saber o porquê.*) 

    ![Escolha o host Oulook desejado: a área de trabalho ou um dos navegadores](images/StartAction.JPG)

6. Pressione F5. Na primeira vez que você fizer isso, você será instruído a especificar o email e a senha do usuário que você usará para depurar o suplemento. Use as credenciais de um administrador para locação do O365. 

    ![Formulário com caixas de texto para email e senha do usuário](images/CredentialsPrompt.JPG)

    >OBSERVAÇÃO: O navegador abrirá a página de logon do Office na Web. (Portanto, se for a primeira vez que você executa o suplemento, insira o nome de usuário e a senha duas vezes.) 

As etapas restantes dependem de você estar executando o suplemento no Outlook ou no Outlook na Web.

### Executar a solução com o Outlook na Web

1. O Outlook para Web será aberto em uma janela do navegador. No Outlook, clique em **Novo** para criar uma nova mensagem de email. 
2. Abaixo do formulário de redação está uma barra de ferramentas com os botões para **Enviar**, **Descartar** e outros utilitários. Dependendo da experiência do **Outlook na Web** que você está usando, o ícone do suplemento estará próximo à extremidade direita dessa barra de ferramentas ou estará no menu suspenso que será aberto quando você clica no botão **...** dessa barra de ferramentas.

   ![Ícone de Suplemento Inserir Arquivos](images/Onedrive_Charts_icon_16x16px.png)

3. Clique no ícone para abrir o suplemento do painel de tarefas.
4. Use o suplemento para adicionar os nomes das três primeiras pastas de trabalho da conta do OneDrive do usuário à mensagem. As páginas e os botões do suplemento são autoexplicativos.

## Executar o projeto com a área de trabalho do Outlook

1. A área de trabalho do Outlook será aberta. No Outlook, clique em **Novo Email** para criar uma nova mensagem de email. 
2. Na faixa de opções **Mensagem** do formulário **Mensagem**, há um botão rotulado **Abrir Suplemento** em um grupo chamado **Arquivos do OneDrive**. Clique no botão para abrir o suplemento.
3. Use o suplemento para adicionar os nomes das três primeiras pastas de trabalho da conta do OneDrive do usuário à mensagem. As páginas e os botões do suplemento são autoexplicativos.

## Problemas conhecidos

* O controle giratório do Fabric só aparece brevemente ou nem isso. 
* Se você estiver executando no Internet Explorer, receberá uma mensagem de erro ao tentar fazer login, dizendo que deve colocar `https://localhost:44301` e `https://outlook.office.com` (ou `https://outlook.office365.com`) na mesma zona de segurança. Mas esse erro ocorre mesmo que você tenha feito isso. 

## Perguntas e comentários

Gostaríamos de receber seus comentários sobre o exemplo *Obter as pastas de trabalho do Excel usando o Microsoft Graph e MSAL em um Suplemento do Office*. Você pode enviar seus comentários na seção *Problemas* deste repositório.
Perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Certifique-se de que as suas perguntas estejam marcadas com [office-js], [MicrosoftGraph] e [API].

## Recursos adicionais

* [Documentação do Microsoft Graph](https://learn.microsoft.com/graph/)
* [Documentação de Suplementos do Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Direitos autorais
Direitos autorais (c) 2019 Microsoft Corporation. Todos os direitos reservados.

Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
