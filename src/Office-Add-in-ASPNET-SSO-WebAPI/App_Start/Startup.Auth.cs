using Owin;
using Microsoft.IdentityModel.Tokens;
using System.Configuration;
using Microsoft.Owin.Security.OAuth;
using Microsoft.Owin.Security.Jwt;
using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;

namespace Office_Add_in_ASPNET_SSO_WebAPI
{
	public partial class Startup
	{
		public void ConfigureAuth(IAppBuilder app)
		{
            TokenValidationParameters tvps = new TokenValidationParameters
            {
				ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
                // Microsoft Accounts have an issuer GUID that is different from any organizational tenant GUID,
                // so to support both kinds of accounts, we do not validate the issuer.
                ValidateIssuer = false,
				SaveSigninToken = true
			};

			app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
			{
				AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
			});
		}
	}
}