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
description: "Découvrir comment créer un complément Microsoft Outlook qui se connecte à Microsoft Graph"
---

# Obtenir des classeurs Excel à l’aide de Microsoft Graph et MSAL dans un complément Outlook 

Découvrez comment créer un complément Microsoft Outlook qui se connecte à Microsoft Graph, qui trouve les trois premiers classeurs stockés dans OneDrive Entreprise, qui récupère leurs noms de fichiers et les insère dans un nouveau formulaire de composition de message dans Outlook.

## Fonctionnalités

L'intégration de données à partir de fournisseurs de services en ligne augmente la valeur et l’adoption de vos compléments. Cet exemple de code vous présente comment connecter votre complément Outlook à Microsoft Graph. Utilisez cet exemple de code pour :

* Se connecter à Microsoft Graph à partir d’un complément Office.
* Utilisez la bibliothèque MSAL .NET pour implémenter l’infrastructure d’autorisation OAuth 2.0 dans un complément.
* Utiliser les API REST OneDrive de Microsoft Graph.
* Afficher une boîte de dialogue à l’aide de l’espace de noms de l’interface utilisateur Office.
* Créer un complément en utilisant l’ASP.NET MVC, de MSAL 3.x.x pour .NET et d’Office.js. 

## Produits concernés

-  Outlook sur l'ensemble des plateformes

## Conditions préalables

Pour exécuter cet exemple de code, les éléments suivants sont requis.

* Visual Studio 2019 ou version ultérieure.

* SQL Server Express (s'il n'est pas installé automatiquement sur les versions récentes de Visual Studio.)

* Compte Office 365 que vous pouvez obtenir en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/devprogramsignup) incluant un abonnement gratuit de 1 an à Office 365.

* Au moins trois classeurs Excel stockés sur OneDrive Entreprise dans votre abonnement Office 365.

* De façon facultative, si vous voulez déboguer sur le bureau plutôt qu’Outlook Online : Office sur Windows, version 1809 ou ultérieure.
* [Outils de développement Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un locataire Microsoft Azure. Ce complément requiert Azure Active Directiory (AD). Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation. Un abonnement d’évaluation peut être demandé ici : [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Auteur(s)
---------|----------
Complément Outlook Microsoft Graph ASP.NET | Microsoft

## Historique des versions

Version | Date | Commentaires
---------| -----| --------
1.0 | 8 juillet 2019 | Publication initiale

## Clause d’exclusion

**CE CODE EST FOURNI *EN L’ÉTAT*, SANS GARANTIE D'AUCUNE SORTE, EXPRESSE OU IMPLICITE, Y COMPRIS TOUTE GARANTIE IMPLICITE D'ADAPTATION À UN USAGE PARTICULIER, DE QUALITÉ MARCHANDE ET DE NON-CONTREFAÇON.**

----------

## Générez et exécutez la solution

## Configurer la solution

1. Dans **Visual Studio**, choisissez le projet **Outlook-Add-in-Microsoft-Graph-ASPNETWeb**. Dans **Propriétés**, assurez-vous que **SSL activé** est défini sur True. Vérifiez que l’**URL SSL** utilise le même nom de domaine et le même numéro de port que ceux répertoriés à l’étape suivante.
 
2. Inscrivez votre application à l’aide du [portail de gestion Azure](https://manage.windowsazure.com). **Connectez-vous à l’aide de l’identité d’un administrateur de votre location Office 365 afin de vous assurer que vous travaillez dans un répertoire Azure Active Directory associé à cette location.** Pour savoir comment inscrire votre application, consulter [Inscrire une application sur la Plateforme d’identités Microsoft](https://learn.microsoft.com/graph/auth-register-app-v2). Utilisez les paramètres suivants :

 - URI DE REDIRECTION : https://localhost:44301/AzureADAuth/Authorize
 - TYPE DE COMPTES PRIS EN CHARGE : « Comptes dans cet annuaire organisationnel uniquement »
 - OCTROI IMPLICITE : Ne pas activer les options d’octroi implicite
 - AUTORISATIONS API (Autorisations déléguées, sans autorisations de l’application) : **Files.Read.All** et **User.Read**

	> Remarque : Une fois que vous avez enregistré votre application, copiez l’**ID d’application (client)** et l’**ID d’annuaire (locataire)** sur le panneau **Vue d’ensemble** de l’inscription de l’application dans le portail de gestion Azure. Lorsque vous créez la clé secrète cliente sur le panneau **Certificats et clés secrètes**, copiez-la également. 
	 
3.  Dans web.config, utilisez les valeurs que vous avez copiées à l’étape précédente. Définissez **AAD:ClientID** sur votre ID client, définissez **AAD:ClientSecret** sur votre clé secrète client et définissez **"AAD:O365TenantID"** sur votre ID locataire. 

## Exécutez la solution

1. Ouvrez le fichier de solution Visual Studio. 
2. Cliquez avec le bouton droit sur la solution **Outlook-Add-in-Microsoft-Graph-ASPNET** dans l’**Explorateur de solutions** (pas les nœuds de projet), puis sélectionnez **Définir les projets de démarrage**. Sélectionnez la case d’option **Plusieurs projets de démarrage**. Assurez-vous que le projet se termine par « Web » apparaît en premier.
3. Dans le menu **Générer**, sélectionnez **Nettoyer la solution**. Une fois l’opération terminée, ouvrez de nouveau le menu **Build**, puis sélectionnez **Générer la solution**.
4. Dans l’**Explorateur de solutions**, sélectionnez le nœud de projet **Outlook-Add-in-Microsoft-Graph-ASPNET** (et non le projet dont le nom se termine par « Web »).
5. Dans le volet **Propriétés**, ouvrez la liste déroulante **Action de démarrage** et indiquez si vous souhaitez exécuter le complément dans la version de bureau d’Outlook ou avec Outlook sur le web dans l’un des navigateurs répertoriés. (*Ne choisissez pas Internet Explorer. Pour en savoir plus, consultez les **Problèmes connus** ci-dessous.*) 

    ![Choisissez l’hôte Oulook souhaité : bureau ou l’un des navigateurs](images/StartAction.JPG)

6. Appuyez sur la touche F5. La première fois que vous effectuez cette opération, vous êtes invité à spécifier l’adresse de courrier et le mot de passe de l’utilisateur que vous utilisez pour le débogage du complément. Utilisez les informations d’identification d’un administrateur pour votre client Office 365. 

    ![Formulaire incluant des zones de texte pour l’adresse de courrier et le mot de passe de l’utilisateur](images/CredentialsPrompt.JPG)

    >REMARQUE : Le navigateur s’ouvre sur la page de connexion pour Office sur le Web. (si vous exécutez le complément pour la première fois, vous devez entrer le nom d’utilisateur et le mot de passe à deux reprises). 

Les étapes restantes varient selon que vous utilisez le complément dans la version de bureau d’Outlook ou Outlook sur le web.

### Exécutez la solution avec Outlook sur le web

1. Outlook pour le web s’ouvre dans une fenêtre du navigateur. Dans Outlook, cliquez sur **Nouveau** pour créer un nouveau message. 
2. Sous le formulaire composer, figure une barre d’outils contenant des boutons permettant d'**Envoyer**, d'**Ignorer** et d’autres utilitaires. Selon l'expérience **Outlook sur le web** que vous utilisez, l’icône du complément se trouve à l’extrémité droite de la barre d’outils ou dans le menu déroulant qui s’ouvre lorsque vous cliquez sur le bouton **...** dans cette barre d’outils.

   ![Icône pour le complément Insérer des fichiers](images/Onedrive_Charts_icon_16x16px.png)

3. Cliquez sur l'icône pour ouvrir le complément de volet Office.
4. Utilisez le complément pour ajouter les noms des trois premiers classeurs dans le message du compte OneDrive de l’utilisateur. Les pages et les boutons du complément sont explicites.

## Exécutez le projet avec la version de bureau d’Outlook

1. La version de bureau d’Outlook s’ouvre. Dans Outlook, cliquez sur **Nouveau message** pour créer un nouveau message. 
2. Dans le ruban **Message** du formulaire **Message**, il existe un bouton intitulé **Ouvrir un complément** dans un groupe appelé **Fichiers OneDrive**. Cliquez sur le bouton pour ouvrir le complément.
3. Utilisez le complément pour ajouter les noms des trois premiers classeurs dans le message du compte OneDrive de l’utilisateur. Les pages et les boutons du complément sont explicites.

## Problèmes connus

* Le contrôle bouton de progression de la structure s’affiche brièvement, voire pas du tout. 
* Si vous exécutez dans Internet Explorer, un message d’erreur apparaît lorsque vous tentez de vous connecter, indiquant que vous devez placer `https://localhost:44301` et `https://outlook.office.com` (ou `https://outlook.office365.com`) dans la même zone de sécurité. Cette erreur se produit même si vous l’avez fait. 

## Questions et commentaires

Nous aimerions recevoir vos commentaires relatifs à l’exemple *Obtenir des classeurs Excel à l’aide de Microsoft Graph et MSAL dans un complément Office*. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.
Si vous avez des questions sur le développement d’Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Assurez-vous que vos questions comportent les balises [office-js], [MicrosoftGraph] et [API].

## Ressources supplémentaires

* [Documentation Microsoft Graph](https://learn.microsoft.com/graph/)
* [Documentation pour compléments Office](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 Microsoft Corporation. Tous droits réservés.

Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/auth/Outlook-Add-in-Microsoft-Graph-ASPNET" />
