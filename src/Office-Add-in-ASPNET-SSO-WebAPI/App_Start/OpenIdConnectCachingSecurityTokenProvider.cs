using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin.Security.Jwt;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Office_Add_in_ASPNET_SSO_WebAPI.App_Start
{
    // This class is necessary because the OAuthBearer Middleware does not leverage
    // the OpenID Connect metadata endpoint exposed by the STS by default.
    public class OpenIdConnectCachingSecurityTokenProvider : IIssuerSecurityKeyProvider
    {
        public Microsoft.IdentityModel.Protocols.ConfigurationManager<OpenIdConnectConfiguration> _configManager;
        private string _issuer;
        private IEnumerable<SecurityKey> _keys;
        private readonly string _metadataEndpoint;

        private readonly ReaderWriterLockSlim _synclock = new ReaderWriterLockSlim();

        public OpenIdConnectCachingSecurityTokenProvider(string metadataEndpoint)
        {
            _metadataEndpoint = metadataEndpoint;
            _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(metadataEndpoint, new OpenIdConnectConfigurationRetriever());

            RetrieveMetadata();
        }

        /// <summary>
        /// Gets the issuer the credentials are for.
        /// </summary>
        /// <value>
        /// The issuer the credentials are for.
        /// </value>
        public string Issuer
        {
            get
            {
                RetrieveMetadata();
                _synclock.EnterReadLock();
                try
                {
                    return _issuer;
                }
                finally
                {
                    _synclock.ExitReadLock();
                }
            }
        }

        /// <summary>
        /// Gets all known security keys.
        /// </summary>
        /// <value>
        /// All known security keys.
        /// </value>
        public IEnumerable<SecurityKey> SecurityKeys
        {
            get
            {
                RetrieveMetadata();
                _synclock.EnterReadLock();
                try
                {
                    return _keys;
                }
                finally
                {
                    _synclock.ExitReadLock();
                }
            }
        }

        private void RetrieveMetadata()
        {
            _synclock.EnterWriteLock();
            try
            {
                OpenIdConnectConfiguration config = Task.Run(_configManager.GetConfigurationAsync).Result;
                _issuer = config.Issuer;
                _keys = config.SigningKeys;
            }
            finally
            {
                _synclock.ExitWriteLock();
            }
        }
    }
}


//using Microsoft.IdentityModel.Protocols;
//using Microsoft.Owin.Security.Jwt;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading;
//using System.Threading.Tasks;
//using System.IdentityModel.Tokens;


//using Microsoft.IdentityModel.Protocols;

//namespace Office_Add_in_ASPNET_SSO_WebAPI.App_Start
//{
//    public class OpenIdConnectCachingSecurityTokenProvider : IIssuerSecurityKeyProvider//IIssuerSecurityTokenProvider
//    {
//        public ConfigurationManager<OpenIdConnectConfiguration> _configManager;
//        private string _issuer;
//        private IEnumerable<SecurityKey> _keys;//<SecurityToken> _tokens;
//        private readonly string _metadataEndpoint;

//        private readonly ReaderWriterLockSlim _synclock = new ReaderWriterLockSlim();

//        public OpenIdConnectCachingSecurityTokenProvider(string metadataEndpoint)
//        {
//            _metadataEndpoint = metadataEndpoint;
//            _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(metadataEndpoint);

//            RetrieveMetadata();
//        }

//        /// <summary>
//        /// Gets the issuer the credentials are for.
//        /// </summary>
//        /// <value>
//        /// The issuer the credentials are for.
//        /// </value>
//        public string Issuer
//        {
//            get
//            {
//                RetrieveMetadata();
//                _synclock.EnterReadLock();
//                try
//                {
//                    return _issuer;
//                }
//                finally
//                {
//                    _synclock.ExitReadLock();
//                }
//            }
//        }

//        /// <summary>
//        /// Gets all known security keys //tokens.
//        /// </summary>
//        /// <value>
//        /// All known security keys //tokens.
//        /// </value>
//        public IEnumerable<SecurityKey> SecurityKeys //<SecurityToken> SecurityTokens
//        {
//            get
//            {
//                RetrieveMetadata();
//                _synclock.EnterReadLock();
//                try
//                {
//                    return _keys; //_tokens;
//                }
//                finally
//                {
//                    _synclock.ExitReadLock();
//                }
//            }
//        }

//        private void RetrieveMetadata()
//        {
//            _synclock.EnterWriteLock();
//            try
//            {
//                OpenIdConnectConfiguration config = _configManager.GetConfigurationAsync().Result;
//                _issuer = config.Issuer;
//                _tokens = config.SigningTokens;
//            }
//            finally
//            {
//                _synclock.ExitWriteLock();
//            }
//        }
//    }
//}