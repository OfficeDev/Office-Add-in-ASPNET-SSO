---
page_type: sample
products:
- office-excel
- office-powerpoint
- office-word
- office-365
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
  - Office 365
  createdDate: 5/1/2017 2:09:09 PM
---
# Office Add-in that that supports Single Sign-on to Office, the Add-in, and Microsoft Graph

The `getAccessToken` API in Office.js enables users who are signed into Office to get access to an AAD-protected add-in and to Microsoft Graph without needing to sign-in again. This sample is built on ASP.NET and Microsoft Identity Library (MSAL) .NET. 

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
* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365. During the preview phase, the SSO requires Office 365 (the subscription version of Office, also called “Click to Run”). You should use the latest monthly version and build from the Insiders channel. You need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). 
    > Note: When a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.
* At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.
* A Microsoft Azure Tenant. This add-in requires Azure Active Directiory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

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
2.0 | October 1, 2019 | Added Display Dialog API fall back.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

----------

## To use the project

### Register the add-in
 
1. Register your application using the [Azure Management Portal](https://manage.windowsazure.com). **Log in with the identity of an administrator of your Office 365 tenancy to ensure that you are working in an Azure Active Directory that is associated with that tenancy.** To learn how to register your application, see [Register an application with the Microsoft Identity Platform](https://docs.microsoft.com/graph/auth-register-app-v2). Use the following settings:

 - NAME: Office-Add-in-ASPNET-SSO
 - REDIRCT URI: https://localhost:44355/AzureADAuth/Authorize
 - SUPPORTED ACCOUNT TYPES: "Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)"
 - IMPLICIT GRANT: Do not enable any Implicit Grant options
 - API PERMISSIONS (Delegated permissions, not Application permissions):

     - **Files.Read.All**
     - **offline_access**
     - **openid**
     - **profile**

  > Note: After you register your application, copy the **Application (client) ID** and the **Directory (tenant) ID** on the **Overview** blade of the App Registration in the Azure Management Portal. When you create the client secret on the **Certificates & secrets** blade, copy it too. 
	 
2. Still in the Azure App registration portal, when you've completed the preceding parts of the registration, select **Expose an API** under **Manage**. Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**. Insert `localhost:44355/` (note the a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. 

3. Select the **Add a scope** button. In the panel that opens, enter `access_as_user` as the **Scope** name.

4. Set **Who can consent?** to **Admins and users**.

5. Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user. Suggestions:

    - **Admin consent title**: Office can act as the user.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.
    - **User consent title**: Office can act as you.
    - **Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.

6. Ensure that **State** is set to **Enabled**.

7. Select **Add scope** .

  > Note: The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

8. In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)

    For each ID, take these steps:

    a. Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Select **Add application**.

10. On the **API permissions** tab, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.

### Configure the solution

1. In **Visual Studio**, right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.

2. Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**. Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first. 

3. Still in **Solution Explorer**, choose the **Office-Add-in-Microsoft-Graph-ASPNETWebAPI** project. In **Properties**, ensure **SSL Enabled** is **True**. Verify that the **SSL URL** is `http://localhost:44355/`.

4. In web.config, use the values that you copied in earlier. Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret. 

	> Note: The **Application (client) ID** is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.

5. Still in the web.config, replace the placeholder `tenant_GUID_here` in the **ida:Issuer**  value with the **Directory (tenant) ID** that you copied when you registered the add-in.

6. If for any reason you set the set SUPPORTED ACCOUNT TYPES to only accounts in your own organization when you registered the add-in, then you must also set **ida:TenantId** in the web.config to the **Directory (tenant) ID** that you copied when you registered the add-in. 

7. In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file. Just above the end `</VersionOverrides>` tag, you'll find the following markup:

    ```
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

8. Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in. The "$" signs are not part of the ID, so do not include them. This is the same ID you used in for the ClientID and Audience in the web.config.

	> Note:  The **Resource** value is the **Application ID URI** you set when you registered the add-in. The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.

### Run the solution

1. Open the Visual Studio solution file. 
2. Right-click **Office-Add-in-ASPNET-SSO** solution in **Solution Explorer** (not the project nodes), and then choose **Set startup projects**. Select the **Multiple startup projects** radio button. Make sure the project that ends with "WebAPI" is listed first.
3. On the **Build** menu, select **Clean Solution**. When it finishes, open the **Build** menu again and select **Build Solution**.
4. In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").
5. In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).

    ![Choose the desired Office host application: Excel or PowerPoint or Word](images/SelectHost.JPG)

6. Press F5. 
7. In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.
8. Click the **Get OneDrive File Names** button. If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane. If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in. After you log in, the file and folder names appear.

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
