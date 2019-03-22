---
topic: sample
products:
- Excel
- PowerPoint
- Word
- Office 365
languages:
- JavaScript
- ASP.NET
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Office 365
  createdDate: 5/1/2017 2:09:09 PM
---
# <a name="office-add-in-that-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Complément Office qui prend en charge l’authentification unique pour Office, le complément et Microsoft Graph

L’API `getAccessTokenAsync` dans Office.js permet aux utilisateurs qui sont connectés à Office d’accéder à un complément protégé par AAD et à Microsoft Graph sans avoir à se reconnecter. Cet exemple repose sur ASP.NET et la bibliothèque MSAL. 

 > Remarque : Cette API`getAccessTokenAsync` est disponible en aperçu.

## <a name="table-of-contents"></a>Sommaire
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Utiliser l’explorateur de projets](#to-use-the-project)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="change-history"></a>Historique des modifications

* 10 mai 2017 : Version d’origine.
* 15 septembre 2017 : Prise en charge de 2FA.
* 8 décembre 2017 : Ajout de la gestion étendue des erreurs.
* 7 janvier 201 : Ajout d’informations sur les pratiques de sécurité d’application web.

## <a name="prerequisites"></a>Conditions requises

* Un compte Office 365.
* L’authentification unique SSO requiert Office 365 (la version par abonnement d’Office, également appelée « Démarrer en un clic »). Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.
* Visual Studio 2017 version 15.4.0 (Aperçu 1) ou version ultérieure

## <a name="deviations-from-best-practices"></a>Écarts entre les meilleures pratiques

Les exemples dans cette repo sont étroitement axées sur la démonstration de l’utilisation de l’API d’authentification unique SSO. Pour effectuer une opération simple, certaines pratiques recommandées ne sont pas suivies, y compris les meilleures pratiques de sécurité de l’application web. *Vous ne devez pas utiliser un de ces exemples comme point de départ du complément production, sauf si vous êtes prêt à apporter des modifications substantielles.* Nous vous recommandons de commencer un complément de production en utilisant l’un des projets de complément Office dans Visual Studio ou en générer un nouveau projet avec la [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office).

Voici l’_un_ des points à retenir concernant ces exemples :

* Les exemples envoient un paramètre de requête codée en dur dans l’URL pour le Microsoft Graph l’API REST. Si vous modifiez ce code dans un complément production et une partie quelconque de paramètre de requête provient d’une intervention de l’utilisateur, n’oubliez pas qu’il est purgé afin qu’il ne puisse pas être utilisé dans une attaque par injection d’en-tête de réponse.

## <a name="to-use-the-project"></a>Utiliser le projet

Cet exemple est destiné à accompagner cette procédure : [Créer un complément Office ASP.NET qui utilise l’authentification unique (aperçu)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-aspnet)

Il existe deux versions de l’exemple dans les dossiers Précédent et Terminé.

Pour utiliser la version précédente et ajouter manuellement le code de l’authentification unique essentiel, suivez les procédures décrites dans l’article en lien ci-dessus.

Pour utiliser avec la version terminée, suivez toutes les procédures, sauf les sections « Code côté client » et « Code côté serveur » dans l’article en lien ci-dessus.

## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle comporte les balises [office-js] et [API].

## <a name="additional-resources"></a>Ressources supplémentaires

* [Documentation de complément Office](https://msdn.microsoft.com/fr-fr/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

Ce projet a adopté le [code de conduite Microsoft Open Source](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

## <a name="copyright"></a>Copyright
Copyright (c) 2017 Microsoft Corporation. Tous droits réservés.

