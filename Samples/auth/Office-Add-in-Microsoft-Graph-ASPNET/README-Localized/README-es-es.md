---
page_type: sample
products:
  - office-excel
languages:
  - javascript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: 5/1/2019 1:25:00 PM
---
# Obtener datos de OneDrive usando Microsoft Graph y MSAL.NET en un complemento de Office 

Aprenda a crear un complemento de Microsoft Office que se conecta a Microsoft Graph, encuentra los tres primeros libros de trabajo almacenados en OneDrive para empresas, obtiene sus nombres de archivo e inserta los nombres en un documento de Office utilizando Office.js.

## Características
La integración de los datos de los proveedores de servicios en línea aumenta el valor y la adopción de sus complementos. En este ejemplo de código se muestra cómo conectar el complemento con Microsoft Graph. Use este ejemplo de código para:

* Conéctese a Microsoft Graph desde un complemento de Office.
* Utilice la biblioteca de MSAL.NET para implementar el marco de autorización de OAuth 2.0 en un complemento.
* Utilice las APIs OneDrive REST de Microsoft Graph.
* Mostrar un diálogo usando el espacio de nombres de la interfaz de usuario de Office.
* Construya un complemento usando ASP.NET MVC, MSAL 3.x.x para .NET, y Office.js. 
* Usar los comandos de un complemento en un complemento

## Se aplica a

-  Excel en Windows (compra única y suscripción)
-  PowerPoint en Windows (compra única y suscripción)
-  Word en Windows (compra única y suscripción)

## Requisitos previos

Para ejecutar este ejemplo de código, se requiere lo siguiente.

* Visual Studio 2019 o posterior.

* SQL Server Express (ya no se instala automáticamente con versiones recientes de Visual Studio).

* Una cuenta de Office 365 que puede obtener al unirse al [programa de desarrollo de Office 365](https://aka.ms/devprogramsignup) que incluye una suscripción gratuita de 1 año a Office 365.

* Al menos tres cuadernos de Excel almacenados en OneDrive para empresas en su suscripción a Office 365.

* Office en Windows, versión 16.0.6769.2001 o superior.

* [Herramientas para desarrolladores de Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un inquilino de Microsoft Azure. Este complemento requiere Azure Active Directiory (AD).  Azure (AD) le ofrece servicios de identidad que las aplicaciones usan para autenticación y autorización. Las suscripciones de prueba se pueden adquirir aquí: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solución

Solución | Autor(es)
---------|-----------
complementos de Office en Microsoft Graph ASP.NET | Microsoft

## Historial de versiones

Versión | Fecha | Comentarios
---------| -----| --------
1.0 |8 de julio de 2019| Lanzamiento inicial

## Renuncia

**ESTE CÓDIGO SE PROPORCIONA*TAL CUAL* SIN GARANTÍA DE NINGÚN TIPO, YA SEA EXPRESA O IMPLÍCITA, INCLUYENDO CUALQUIER GARANTÍA IMPLÍCITA DE IDONEIDAD PARA UN PROPÓSITO PARTICULAR, COMERCIABILIDAD O NO INFRACCIÓN. **

----------

## Compilar y ejecutar la solución

### Configurar la solución

1. En **Visual Studio**, elija el proyecto**Office-Add-in-Microsoft-Graph-ASPNETWeb**. En **Propiedades**, asegúrese de que el**SSL esté activado** y sea **Verdadero**. Compruebe que la **URL de SSL** use el mismo nombre de dominio y número de puerto que se indica en el paso 3 que se muestra a continuación.
 
2. Registre la aplicación mediante el [Portal de administración de Azure](https://manage.windowsazure.com). **Ingrese con la identidad de un administrador de su Oficina 365 para asegurarse de que está trabajando en un Directorio Activo Azure que está asociado con esa tenencia.** Para saber cómo registrar aplicaciones, consulte [Registrar una aplicación en el Microsoft Identity Platform](https://learn.microsoft.com/graph/auth-register-app-v2). Use la siguiente configuración:

 - URI REDIRCT: https://localhost:44301/AzureADAuth/Authorize
 - TIPOS DE CUENTA ADMITIDAS: «Solo las cuentas de este directorio organizativo»
 - CONCESIÓN IMPLÍCITA: No habilitar ninguna opción de subvención implícita
 - PERMISOS DE LA API (Permisos delegados, no permisos de aplicación): **Files.Read.All** y **User.Read**

	> Nota: Después de registrar la aplicación, copie la **Id. de la aplicación (cliente)** y el**Id. del directorio (inquilino)** en la hoja de **información general** del registro de la aplicación en el Portal de administración de Azure. Cuando cree el secreto de cliente en la hoja de **Certificados y Secretos**, cópielo. 
	 
3.  En web.config, use los valores que copió en el paso anterior. Establezca **AAD:ClientID** en su identificación de cliente, **AAD:ClientSecret** en el secreto de cliente, y finalmente **"AAD:O365TenantID"** en la identificación de inquilino  

### Ejecute la solución

1. Abra el archivo de la solución de Visual Studio. 
2. Haga clic con el botón derecho en solución en**Office-Add-in-Microsoft-Graph-ASPNET** en el **Explorador de soluciones ** (no en los nodos del proyecto) y luego, seleccione **establecer proyectos de inicio**. Seleccione el botón de radio **Proyectos de inicio múltiples**. Asegúrate de que el proyecto que termina con "Web" aparece en primer lugar.
3. En el menú **compilación**, seleccione **Limpiar solución**. Cuando termine, abra de nuevo el menú **Compilación**. y seleccione **Solución de compilación**.
4. En el **Explorador de soluciones**, seleccione el nodo de proyecto **Office-Add-in-ASPNET-SSO**nodo del proyecto (no el nodo superior de la solución y no el proyecto cuyo nombre termina en "Web").
5. En el panel** Propiedades**, abra la lista desplegable **niciar documento** y elija una de las tres opciones (Excel, Word o PowerPoint).

    ![ Elija la aplicación host de Office que desee:](images/SelectHost.JPG) Word, Excel o PowerPoint](images/SelectHost.JPG)

6. Pulse <kbd>F5</kbd>. 
7. En la aplicación de Office, elija **insertar** > **Abrir complemento**en los**archivos de OneDrive** para abrir el complemento del panel de tareas.
8. Las páginas y los botones del complemento se explican por sí mismos. 

## Problemas conocidos

* El control del hilandero de la tela aparece sólo brevemente o no aparece en absoluto.

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio.
Las preguntas sobre el desarrollo de complementos de oficina deben enviarse a [Stack Overflow](http://stackoverflow.com). Asegúrate de que tus preguntas estén etiquetadas con [office-js] y [MicrosoftGraph].

## Recursos adicionales

* [Documentación de Microsoft Graph](https://learn.microsoft.com/graph/)
* [Documentación de complementos de Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Derechos de autor
Derechos de autor (c) 2019 Microsoft Corporation. Todos los derechos reservados.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte[Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
