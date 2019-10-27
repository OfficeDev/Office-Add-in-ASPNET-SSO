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
using System.Web.Http;
using System;
using System.Net;
using System.Net.Http;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Controllers
{
	[Authorize]
    public class ValuesController : ApiController
    {
		// GET api/values
		public async Task<HttpResponseMessage> Get()
		{
            // OWIN middleware validated the audience and issuer, but the scope must also be validated; must contain "access_as_user".
            string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
            if (!(addinScopes.Contains("access_as_user")))
            {
                return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
            }

            // Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
            // Beginning with MSAL.NET 3.x.x, the bootstrapContext is just the bootstrap token itself.
            string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString(); 
            UserAssertion userAssertion = new UserAssertion(bootstrapContext);

			var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                          .WithRedirectUri("https://localhost:44355")
                                                          .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                          .Build();

            // MSAL.NET adds the profile, offline_access, and openid scopes itself. It will throw an error if you add
            // them redundantly here.
			string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };

			// Get the access token for Microsoft Graph.
			AcquireTokenOnBehalfOfParameterBuilder result = null;
			try
			{
				result = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
			}
			catch (MsalServiceException e)
			{
				// Handle request for multi-factor authentication.
				if (e.Message.StartsWith("AADSTS50076"))
				{
					string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
					return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
                    // The client should recall the getAccessToken function and pass the claims string as the 
                    // authChallenge value in the function's Options parameter.
				}

				// Handle invalid scope (permission).
				if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
				{
					return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
				}

				// Handle all other MsalServiceExceptions.
				else
				{
					throw e;
				}
			}

            AuthenticationResult authResult;
            try
            {
                authResult = await result.ExecuteAsync();
            }
            catch (Exception e)
            {
                return HttpErrorHelper.SendErrorToClient(HttpStatusCode.BadRequest, e, null);
            }

            return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
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
