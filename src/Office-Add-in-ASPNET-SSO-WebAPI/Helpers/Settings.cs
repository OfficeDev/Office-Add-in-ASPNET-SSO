// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System;
using System.Configuration;
using System.Web;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Helpers
{
	/// <summary>
	/// Provides management of basic user and web application authentication and authorization information. 
	/// </summary>
	public static class Settings
	{
        static Settings()
        {
            if (String.IsNullOrEmpty(ConfigurationManager.AppSettings["ida:TenantId"])
                || ConfigurationManager.AppSettings["ida:TenantId"] == "{Tenant GUID}")
            {
                // Accounts in any organization and possibly also Microsoft Accounts can sign-in.
                AzureADAuthority = @"https://login.microsoftonline.com/common/oauth2/v2.0";
            }
            else
            {
                // Only accounts in a specific tenancy are allowed to sign in.
                AzureADAuthority = @"https://login.microsoftonline.com/" + ConfigurationManager.AppSettings["ida:TenantId"] + "/oauth2/v2.0";
            }
        }


        public static string AzureADClientId = ConfigurationManager.AppSettings["ida:ClientID"];
		public static string AzureADClientSecret = ConfigurationManager.AppSettings["ida:Password"];
		public static string AzureADLogoutAuthority = @"https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=";
		public static string GraphApiResource = @"https://graph.microsoft.com/";
        public static string AzureADAuthority; 
    }
}
