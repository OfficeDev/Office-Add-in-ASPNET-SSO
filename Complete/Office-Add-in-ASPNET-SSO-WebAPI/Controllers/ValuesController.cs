// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the repo.

/* 
    This file provides controller methods to get data from MS Graph. 
*/

using Microsoft.Identity.Client;
using System.IdentityModel.Tokens;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
using System;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Controllers
{
    [Authorize]
    public class ValuesController : ApiController
    {
        // GET api/values
        public async Task<IEnumerable<string>> Get()
        {
            // OWIN middleware validated the audience and issuer, but the scope must also be validated; must contain "access_as_user".
            string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
            if (addinScopes.Contains("access_as_user"))
            {
                // Get the raw token that the add-in page received from the Office host.
                var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext
                    as BootstrapContext;
                UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);

                // Get the access token for MS Graph. 
                ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
                ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
                string[] graphScopes = { "Files.Read.All" };                
                AuthenticationResult result = null;
                try
                {
                    // The AcquireTokenOnBehalfOfAsync method will first look in the MSAL in memory cache for a
                    // matching access token. Only if there isn't one, does it initiate the "on behalf of" flow
                    // with the Azure AD V2 endpoint.
                    result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
                }
                catch (MsalUiRequiredException e)
                {
                    // If multi-factor authentication is required by the MS Graph resource an
                    // the user has not yet provided it, AAD will throw an exception containing a 
                    // Claims property.
                    if (String.IsNullOrEmpty(e.Claims))
                    {
                        throw e;
                    }
                    else
                    {
                        // The Claims property value must be passed to the client which will pass it
                        // to the Office host, which will then include it in a request for a new token.
                        // AAD will prompt the user for all required forms of authentication.
                        throw new HttpException(e.Claims);
                    }   
                }

                // Get the names of files and folders in OneDrive for Business by using the Microsoft Graph API. Select only properties needed.
                var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
                var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);

                // The returned JSON includes OData metadata and eTags that the add-in does not use. 
                // Return to the client-side only the filenames.
                List<string> itemNames = new List<string>();
                foreach (OneDriveItem item in getFilesResult)
                {
                    itemNames.Add(item.Name);
                }
                return itemNames;
            }
            return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
