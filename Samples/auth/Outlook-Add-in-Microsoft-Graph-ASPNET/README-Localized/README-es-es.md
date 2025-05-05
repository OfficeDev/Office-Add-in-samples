---
page_type: sample
products:
  - m365
  - office
  - office-outlook
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
description: "Obtenga más información sobre cómo crear un complemento de Microsoft Outlook que se conecte a Microsoft Graph"
---

# Obtenga libros de Excel cuando usa Microsoft Graph y MSAL en un complemento de Outlook. 

Infórmese sobre cómo crear un complemento de Microsoft Outlook que se conecta a Microsoft Graph, busca los primeros tres libros almacenados en OneDrive para la Empresa, encuentra sus nombres de archivo y los inserta en un nuevo formulario de redacción de mensajes en Outlook.

## Características

Integrar datos de proveedores de servicios en línea aumenta el valor y la adopción de los complementos. En este ejemplo de código se muestra cómo conectar el complemento de Outlook con Microsoft Graph. Use este ejemplo de código para:

* Conectarse a Microsoft Graph desde un complemento de Office.
* Use la biblioteca MSAL .NET para implementar el marco de autorización OAuth 2.0 en un complemento.
* Usar las API de REST de OneDrive y Excel desde Microsoft Graph.
* Mostrar un diálogo usando el espacio de nombres de la interfaz de usuario de Office.
* Construir un complemento usando ASP.NET MVC, MSAL 3.x.x para .NET, y Office.js. 

## Se aplica a

-  Outlook en todas las plataformas

## Requisitos previos

Para ejecutar este ejemplo de código, se requiere lo siguiente.

* Visual Studio 2019 o posterior.

* SQL Server Express (ya no se instala automáticamente con versiones recientes de Visual Studio).

* Una cuenta de Office 365, la cual puede obtener al unirse al [programa para desarrolladores de Office 365](https://aka.ms/devprogramsignup), que incluye una suscripción gratuita de 1 año a Office 365.

* Al menos tres libros de Excel almacenados en OneDrive para la Empresa en su suscripción a Office 365.

* Opcional, si desea depurar en el escritorio en lugar de en Outlook Online: Outlook para Windows, versión 1809 o superior.
* [Herramientas para desarrolladores de Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un inquilino de Microsoft Azure. Este complemento requiere Azure Active Directiory (AD).  Azure (AD) le ofrece servicios de identidad que las aplicaciones usan para autenticación y autorización. Las suscripciones de prueba se pueden adquirir aquí: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solución

Solución | Autor(es)
---------|-----------
complementos de Outlook en Microsoft Graph ASP.NET | Microsoft

## Historial de versiones

Versión | Fecha | Comentarios
---------| -----| --------
1.0 |8 de julio de 2019| Lanzamiento inicial

## Renuncia

**ESTE CÓDIGO SE PROPORCIONA*TAL CUAL* SIN GARANTÍA DE NINGÚN TIPO, YA SEA EXPRESA O IMPLÍCITA, INCLUYENDO CUALQUIER GARANTÍA IMPLÍCITA DE IDONEIDAD PARA UN PROPÓSITO PARTICULAR, COMERCIABILIDAD O NO INFRACCIÓN. **

----------

## Compilar y ejecutar la solución

## Configurar la solución

1. En **Visual Studio**, elija el proyecto**Outlook-Add-in-Microsoft-Graph-ASPNETWeb**. Asegúrese de que en las **propiedades** el **SSL esté activado** y sea **verdadero**. Compruebe que la **URL de SSL** use el mismo nombre de dominio y número de puerto que se indica en el paso 3 que se muestra a continuación.
 
2. Registre la aplicación mediante el [Portal de administración de Azure](https://manage.windowsazure.com). **Ingrese con la identidad de un administrador de su Oficina 365 para asegurarse de que está trabajando en un Directorio Activo Azure que está asociado con esa tenencia.** Para saber cómo registrar aplicaciones, consulte [Registrar una aplicación en el Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2). Use la siguiente configuración:

 - URI REDIRCT: https://localhost:44301/AzureADAuth/Authorize
 - TIPOS DE CUENTA ADMITIDAS: «Solo las cuentas de este directorio organizativo»
 - CONCESIÓN IMPLÍCITA: No habilitar ninguna opción de subvención implícita
 - PERMISOS DE LA API (Permisos delegados, no permisos de aplicación): **Files.Read.All** y **User.Read**

	> Nota: Después de registrar la aplicación, copie la **Id. de la aplicación (cliente)** y el**Id. del directorio (inquilino)** en la hoja de **información general** del registro de la aplicación en el Portal de administración de Azure. Cuando cree el secreto de cliente en la hoja de **Certificados y Secretos**, cópielo. 
	 
3.  En web.config, use los valores que copió en el paso anterior. Establezca **AAD:ClientID** en su identificación de cliente, **AAD:ClientSecret** en el secreto de cliente, y finalmente **"AAD:O365TenantID"** en la identificación de inquilino  

## Ejecute la solución

1. Abra el archivo de la solución de Visual Studio. 
2. Haga clic con el botón derecho en solución en**Outlook-Add-in-Microsoft-Graph-ASPNET** en el **Explorador de soluciones ** (no en los nodos del proyecto) y luego, seleccione **establecer proyectos de inicio**. Seleccione el botón de selección **Proyectos de inicio múltiples**. Asegúrate de que el proyecto que termina con "Web" aparece en primer lugar.
3. En el menú **compilación**, seleccione **Limpiar solución**. Cuando termine, abra de nuevo el menú **Compilación** y seleccione **Solución de compilación**.
4. En el **Explorador de soluciones**, seleccione el nodo de proyecto **Outlook-Add-in-Microsoft-Graph-ASPNET**nodo del proyecto (no el nodo superior de la solución y no el proyecto cuyo nombre termina en "Web").
5. En el panel **propiedades**, abra en el menú desplegable **iniciar acción** y elija si desea ejecutar el complemento en el escritorio de Outlook o con Outlook en la web en uno de los navegadores de la lista. (*No elija Internet Explorer. Vea los siguientes **problemas conocidos** para saber por qué.*) 

    ![Elija el servidor de Oulook deseado: escritorio o uno de los exploradores](images/StartAction.JPG)

6. Pulse F5. La primera vez que haga esto, se le pedirá que especifique el correo electrónico y la contraseña del usuario que usará para depurar el complemento. Use las credenciales del administrador en el espacio empresarial de O365. 

    ![Formulario con cuadros de texto para el correo electrónico y la contraseña del usuario](images/CredentialsPrompt.JPG)

    >NOTA: Se abrirá el explorador en la página de inicio de sesión de Office en la web. (Por lo tanto, si esta es la primera vez que ejecuta el complemento, deberá escribir el nombre de usuario y la contraseña dos veces.) 

Los pasos restantes dependen de si está ejecutando el complemento en el escritorio de Outlook u Outlook en la Web.

### Ejecute la solución con Outlook en la Web

1. Outlook para web se abrirá en una ventana del explorador. En Outlook, haga clic en **nuevo** para crear un mensaje de correo electrónico. 
2. Debajo del formulario de redacción está una barra de herramientas con botones para **enviar**, **descartar**, entre otras utilidades. En función de la experiencia de **Outlook en la Web** que use, el icono del complemento estará cerca del extremo derecho de la barra de herramientas, o bien en el menú desplegable que se abre al hacer clic en el botón **...** de la barra de herramientas.

   ![Icono del complemento para insertar archivos](images/Onedrive_Charts_icon_16x16px.png)

3. Haga clic en el icono para abrir el complemento del panel de tareas.
4. Use el complemento para agregar al mensaje los nombres de los tres primeros libros de la cuenta OneDrive del usuario. Las páginas y los botones del complemento se explican por sí mismos.

## Ejecutar el proyecto con la versión de escritorio de Outlook

1. Se abrirá la versión de escritorio de Outlook. En Outlook, haga clic en **nuevo correo** para crear un mensaje de correo electrónico. 
2. En la cinta de **mensajes** del formulario de **mensajes**, hay un botón etiquetado como **complemento abierto** en un grupo llamado **archivos de OneDrive**. Haga clic en el botón para abrir el complemento.
3. Use el complemento para agregar al mensaje los nombres de los tres primeros libros de la cuenta OneDrive del usuario. Las páginas y los botones del complemento se explican por sí mismos.

## Problemas conocidos

* El control del hilandero de la tela aparece sólo brevemente o no aparece en absoluto. 
* Si lo está ejecutando en Internet Explorer, recibirá un mensaje de error cuando intente iniciar sesión, donde se indicará que deberá colocar `https://localhost:44301` y `https://outlook.office.com` (u `https://outlook.office365.com`) en la misma zona de seguridad. Pero, este error se produce incluso después de haberlo completado. 

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre la muestra para *obtener libros de Excel usando Microsoft Graph y MSAL en un complemento de Office* Puede enviarnos comentarios a través de la sección de *problemas* del repositorio.
Las preguntas generales sobre desarrollo en Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrese de que sus preguntas están etiquetadas con [office-js], [MicrosoftGraph] y [API].

## Recursos adicionales

* [Documentación de Microsoft Graph](https://learn.microsoft.com/graph/)
* [Documentación de complementos de Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Derechos de autor
Copyright (c) 2019 Microsoft Corporation. Todos los derechos reservados.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
