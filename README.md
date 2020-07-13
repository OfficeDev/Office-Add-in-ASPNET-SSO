---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-word
- microsoft-365
languages:
- javascript
- aspx
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Microsoft 365
  createdDate: 5/1/2017 2:09:09 PM
---
# Office Add-in that that supports Single Sign-on to Office, the Add-in, and Microsoft Graph

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. This sample is built on ASP.NET and Microsoft Identity Library (MSAL) .NET.

There are two versions of the sample in this repo, one of which has its own README file:

- In the **Begin** folder is the starting point for the SSO walkthrough at at [Create an ASP.NET Office Add-in that uses single sign-on](https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet). Please follow the instructions in the article.
- In the **Complete** folder is the completed sample you would have if you completed the walkthrough. To use this version, follow the instructions in the article [Create an ASP.NET Office Add-in that uses single sign-on](https://docs.microsoft.com/office/dev/add-ins/develop/create-sso-office-add-ins-aspnet), but substitute "Complete" for "Begin" in those instructions and skip the sections **Code the client-side** and **Code the server-side**.

 > Note: The `getAccessToken` API is in preview.

## Features

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* See how to use the Single Sign-on (SSO) API
* Connect to Microsoft Graph from an Office Add-in.
* Build an Add-in using ASP.NET MVC, MSAL 3.x.x for .NET, and Office.js. 
* Use the MSAL.NET Library to implement the OAuth 2.0 authorization framework in an add-in.
* Use the OneDrive REST APIs from Microsoft Graph.
* See how an add-in can fall back to an interactive sign-in in scenarios where SSO is not supported.
* Show a dialog using the Office UI namespace in scenarios where SSO is not supported.
* Use add-in commands in an add-in.

## Applies to

-  Excel on Windows (subscription)
-  PowerPoint on Windows (subscription)
-  Word on Windows (subscription)

## Prerequisites

To run this code sample, the following are required.

* Visual Studio 2019 or later.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* An Microsoft 365 account which you can get by joining the [Microsoft 365 developer program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Microsoft 365. During the preview phase, the SSO requires a Microsoft 365 (which includes the subscription version of Office). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). 
    > Note: When a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
* At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.
* A Microsoft Azure Tenant. This add-in requires Azure Active Directory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Solution

Solution | Author(s)
---------|----------
Office Add-in Microsoft Graph ASP.NET | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | May 10, 2017| Initial release
1.0 | September 15, 2017 | Added support for 2FA.
1.0 | December 8, 2017 | Added extensive error handling.
1.0 | January 7, 2019 | Added information about web application security practices.
2.0 | November 5, 2019 | Added Display Dialog API fall back and use new version of SSO API.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## Security note

The sample sends a hardcoded query parameter on the URL for the Microsoft Graph REST API. If you modify this code in a production add-in and any part of query parameter comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about developing Office Add-ins should be posted to [Stack Overflow](http://stackoverflow.com). Ensure your questions are tagged with [office-js] and [MicrosoftGraph].


## Additional resources

* [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2019 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/auth/Office-Add-in-ASPNET-SSO" />
