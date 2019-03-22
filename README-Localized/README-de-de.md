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
# <a name="office-add-in-that-that-supports-single-sign-on-to-office-the-add-in-and-microsoft-graph"></a>Office-Add-In, das einmaliges Anmelden bei Office, beim Add-In und bei Microsoft Graph unterstützt

Die `getAccessTokenAsync`-API in Office.js ermöglicht, dass Benutzer, die bei Office angemeldet sind, Zugriff auf ein durch AAD geschütztes Add-In und auf Microsoft Graph erhalten, ohne sich erneut anmelden zu müssen. Dieses Beispiel basiert auf ASP.NET und der Microsoft-Identitätsbibliothek (MSAL). 

 > Hinweis: Die `getAccessTokenAsync`-API befindet sich in der Vorschau.

## <a name="table-of-contents"></a>Inhaltsverzeichnis
* [Änderungsverlauf](#change-history)
* [Voraussetzungen](#prerequisites)
* [Verwenden des Projekts](#to-use-the-project)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="change-history"></a>Änderungsverlauf

* 10. Mai 2017: Ursprüngliche Version.
* 15. September 2017: Zusätzliche Unterstützung für 2FA.
* 8. Dezember 2017: Umfassende Fehlerbehandlung hinzugefügt.
* 7. Januar 2019: Informationen zu Schutzmaßnahmen für die Anwendungssicherheit hinzugefügt.

## <a name="prerequisites"></a>Voraussetzungen

* Ein Office 365-Konto.
* Während der Vorschauphase erfordert SSO Office 365 (die Abonnementversion von Office, auch als "Klick-und-Los" bezeichnet). Sie sollten die neueste monatliche Version und den neuesten monatlichen Build aus dem Insider-Kanal verwenden. Sie müssen Office-Insider sein, um diese Version nutzen zu können. Weitere Informationen finden Sie unter [Office-Insider werden](https://products.office.com/office-insider?tab=tab-1). Bitte beachten Sie: Wenn ein Build zum halbjährlichen Produktionskanal hochgestuft wird, ist der Support für Vorschaufeatures, einschließlich SSO, bei diesem Build deaktiviert.
* Visual Studio 2017 Version 15.4.0 Vorschauversion 1 oder höher.

## <a name="deviations-from-best-practices"></a>Abweichungen von bewährten Methoden

Die Beispiele in diesem Repository konzentrieren sich fast ausschließlich auf die Demonstration der Verwendung der SSO-APIs. Um die Beispiele einfach zu halten, werden einige bewährte Methoden nicht befolgt, einschließlich bewährter Methoden zur Sicherheit von Webanwendungen. *Sie sollten keines dieser Beispiele als Ausgangspunkt für ein Produktions-Add-In verwenden, es sei denn, Sie sind bereit, wesentliche Änderungen vorzunehmen.* Wir empfehlen, dass Sie ein Produktions-Add-In beginnen, indem Sie eines der Office-Add-In-Projekte in Visual Studio verwenden oder indem Sie ein neues Projekt mit dem [Yeoman-Generator für Office Add-Ins](https://github.com/OfficeDev/generator-office) erstellen.

_Einer_ Punkte, der im Hinblick auf diese Beispiele beachtet werden muss:

* Die Beispiele senden einen hartcodierten Abfrageparameter für die URL für die Microsoft Graph-REST-API. Wenn Sie diesen Code in einem Produktions-Add-In ändern und ein Teil des Abfrageparameters aus Benutzereingaben stammt, stellen Sie sicher, dass er bereinigt wird, sodass er nicht in einem Angriff mit Antwortheadereinschleusung verwendet werden kann.

## <a name="to-use-the-project"></a>Verwenden des Projekts

Dieses Beispiel soll die folgende exemplarische Vorgehensweise begleiten: [Erstellen eines ASP.NET-Office-Add-Ins, das einmaliges Anmelden verwendet (Vorschau)](https://dev.office.com/docs/add-ins/develop/create-sso-office-add-ins-aspnet)

Es gibt zwei Versionen des Beispiels in den Ordnern "Before" und "Completed".

Um die Before-Version zu verwenden und den entscheidenden SSO-orientierten Code manuell hinzuzufügen, befolgen Sie alle Verfahren im oben verlinkten Artikel.

Um mit der Completed Version zu arbeiten, befolgen Sie alle Verfahren mit Ausnahme der Abschnitte "Codieren der Clientseite" und "Codieren der Serverseite" im oben verlinkten Artikel.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich dieses Beispiels. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden.

Fragen zur Microsoft Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API) gestellt werden. Wenn Ihre Frage die Office JavaScript-APIs betrifft, sollte die Frage mit [office-js] und [API] kategorisiert sein.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* [Dokumentation zu Office-Add-Ins](https://msdn.microsoft.com/de-de/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* Weitere Office-Add-In-Beispiele unter [OfficeDev auf Github](https://github.com/officedev)

In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).

## <a name="copyright"></a>Copyright
Copyright (c) 2017 Microsoft Corporation. Alle Rechte vorbehalten.

