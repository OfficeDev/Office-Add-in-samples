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
# Obtenir des données OneDrive à l’aide de Microsoft Graph et MSAL.NET dans un complément Office 

Découvrez comment créer un complément Microsoft Office qui se connecte à Microsoft Graph, qui trouve les trois premiers classeurs stockés dans OneDrive Entreprise, qui récupère leurs noms de fichiers et les insère dans un document Office à l’aide de Office.js.

## Fonctionnalités
Le fait d’intégrer des données à partir de fournisseurs de services en ligne augmente la valeur et l’adoption de vos compléments. Cet exemple de code vous montre comment connecter votre complément à Microsoft Graph. Utilisez cet exemple de code pour :

* Se connecter à Microsoft Graph à partir d’un complément Office.
* Utiliser la bibliothèque MSAL.NET pour implémenter l’infrastructure d’autorisation OAuth 2.0 dans un complément.
* Utiliser les API REST OneDrive à partir de Microsoft Graph.
* Afficher une boîte de dialogue à l’aide de l’espace de noms de l’interface utilisateur Office.
* Créer un complément à l’aide d’ASP.NET MVC, de MSAL 3.x.x pour .NET et d’Office.js. 
* Utiliser les commandes de complément dans un complément.

## S’applique à

-  Excel sur Windows (achat définitif et abonnement)
-  PowerPoint sur Windows (achat définitif et abonnement)
-  Word sur Windows (achat définitif et abonnement)

## Conditions préalables

Pour exécuter cet exemple de code, les éléments suivants sont requis.

* Visual Studio 2019 ou version ultérieure.

* SQL Server Express (N’est plus installé automatiquement avec les versions récentes de Visual Studio.)

* Compte Office 365 que vous pouvez obtenir en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/devprogramsignup) incluant un abonnement gratuit de 1 an à Office 365.

* Au moins trois classeurs Excel stockés sur OneDrive Entreprise dans votre abonnement Office 365.

* Office sur Windows, version 16.0.6769.2001 ou ultérieure.

* [Outils de développement Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un locataire Microsoft Azure. Ce complément requiert Azure Active Directiory (AD). Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation. Un abonnement d’évaluation peut être demandé ici : [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Auteur(s)
---------|----------
complément Office Microsoft Graph ASP.NET | Microsoft

## Historique des versions

Version | Date | Commentaires
---------| -----| --------
1.0 | 8 juillet 2019 | Publication initiale

## Clause d’exclusion

**CE CODE EST FOURNI *EN L’ÉTAT*, SANS GARANTIE D'AUCUNE SORTE, EXPRESSE OU IMPLICITE, Y COMPRIS TOUTE GARANTIE IMPLICITE D'ADAPTATION À UN USAGE PARTICULIER, DE QUALITÉ MARCHANDE ET DE NON-CONTREFAÇON.**

----------

## Générez et exécutez la solution

### Configurer la solution

1. Dans **Visual Studio**, choisissez le projet **Office-Add-in-Microsoft-Graph-ASPNETWeb**. Dans **Propriétés**, assurez-vous que **SSL activé** est défini sur **True**. Vérifiez que l’**URL SSL** utilise le même nom de domaine et le même numéro de port que ceux répertoriés à l’étape suivante.
 
2. Inscrivez votre application à l’aide du [portail de gestion Azure](https://manage.windowsazure.com). **Connectez-vous à l’aide de l’identité d’un administrateur de votre location Office 365 afin de vous assurer que vous travaillez dans un répertoire Azure Active Directory associé à cette location.** Pour savoir comment inscrire votre application, consulter [Inscrire une application sur la Plateforme d’identités Microsoft](https://docs.microsoft.com/graph/auth-register-app-v2). Utilisez les paramètres suivants :

 - URI DE REDIRECTION : https://localhost:44301/AzureADAuth/Authorize
 - TYPE DE COMPTES PRIS EN CHARGE : « Comptes dans cet annuaire organisationnel uniquement »
 - OCTROI IMPLICITE : Ne pas activer les options d’octroi implicite
 - AUTORISATIONS API (Autorisations déléguées, sans autorisations de l’application) : **Files.Read.All** et **User.Read**

	> Remarque : Une fois que vous avez enregistré votre application, copiez l’**ID d’application (client)** et l’**ID d’annuaire (locataire)** sur le panneau **Vue d’ensemble** de l’inscription de l’application dans le portail de gestion Azure. Lorsque vous créez la clé secrète cliente sur le panneau **Certificats et clés secrètes**, copiez-la également. 
	 
3.  Dans web.config, utilisez les valeurs que vous avez copiées à l’étape précédente. Définissez **AAD:ClientID** sur votre ID client, définissez **AAD:ClientSecret** sur votre clé secrète client et définissez **"AAD:O365TenantID"** sur votre ID locataire. 

### Exécutez la solution

1. Ouvrez le fichier de solution Visual Studio. 
2. Cliquez avec le bouton droit sur solution **Office-Add-in-Microsoft-Graph-ASPNET** dans l’**Explorateur de solutions** (pas les nœuds de projet), puis sélectionnez **définir les projets de démarrage**. Sélectionnez la case d’option **Plusieurs projets de démarrage**. Assurez-vous que le projet se termine par « Web » apparaît en premier.
3. Dans le menu **Générer**, sélectionnez **Nettoyer la solution**. Une fois l’opération terminée, ouvrez de nouveau le menu **Build**, puis sélectionnez **Générer la solution**.
4. Dans l’**Explorateur de solutions**, sélectionnez le nœud de projet **Office-Add-in-Microsoft-Graph-ASPNET** (et non le projet dont le nom se termine par « WebAPI »).
5. Dans le volet **Propriétés**, ouvrez la liste déroulante **Document de départ**, puis choisissez l’une des trois options (Excel, Word ou PowerPoint).

    ![Choisissez l’application hôte Office souhaitée :](images/SelectHost.JPG) Excel ou PowerPoint ou Word](images/SelectHost.JPG)

6. Appuyez sur la touche F5. 
7. Dans l’application Office, sélectionnez **Insérer** > **Ouvrir un complément** dans le groupe **Fichiers OneDrive** pour ouvrir le complément du volet Office.
8. Les pages et les boutons du complément sont explicites. 

## Problèmes connus

* Le contrôle bouton fléché Fabric s’affiche brièvement, voire pas du tout.

## Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.
Si vous avez des questions sur le développement des compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com). Assurez-vous que vos questions comportent les balises [office-js] et [MicrosoftGraph].

## Ressources supplémentaires

* [Documentation Microsoft Graph](https://docs.microsoft.com/graph/)
* [Documentation pour compléments Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 Microsoft Corporation. Tous droits réservés.

Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Office-Add-in-Microsoft-Graph-ASPNET" />
