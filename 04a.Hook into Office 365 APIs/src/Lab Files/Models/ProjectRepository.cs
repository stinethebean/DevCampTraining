using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Exchange;
using Microsoft.Office365.OAuth;
using Microsoft.Office365.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace MVCResearchTracker.Models
{
    public class ProjectRepository
    {
        private string SiteUrl = "https://[tenancy].sharepoint.com";
        public async Task<List<MyContact>> GetMyContacts()
        {
             //code here
        }
        private async Task<ExchangeClient> EnsureExchangeClientCreated()
        {

            DiscoveryContext disco = GetFromCache("DiscoveryContext") as DiscoveryContext;

            if (disco == null)
            {
                disco = await DiscoveryContext.CreateAsync();
                SaveInCache("DiscoveryContext", disco);
            }

            string ServiceResourceId = "https://outlook.office365.com";
            Uri ServiceEndpointUri = new Uri("https://outlook.office365.com/ews/odata");


            var dcr = await disco.DiscoverResourceAsync(ServiceResourceId);

            string clientId = disco.AppIdentity.ClientId;
            string clientSecret = disco.AppIdentity.ClientSecret;
            SaveInCache("LastLoggedInUser", dcr.UserId);


            ExchangeClient exClient = new ExchangeClient(ServiceEndpointUri, async () =>
            {
                ClientCredential creds = new ClientCredential(clientId, clientSecret);
                UserIdentifier userId = new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId);
                AuthenticationContext authContext = disco.AuthenticationContext;

                AuthenticationResult authResult = await authContext.AcquireTokenSilentAsync(ServiceResourceId, creds, userId);

                return authResult.AccessToken;
            });
            disco.ClearCache();
            return exClient;
        }
        public async Task<List<MyFile>> GetMyFiles()
        {
            //code here
        }
        private async Task<SharePointClient> EnsureFilesClientCreated()
        {
            DiscoveryContext disco = GetFromCache("DiscoveryContext") as DiscoveryContext;

            if (disco == null)
            {
                disco = await DiscoveryContext.CreateAsync();
                SaveInCache("DiscoveryContext", disco);
            }

            var dcr = await disco.DiscoverCapabilityAsync("MyFiles");

            var ServiceResourceId = dcr.ServiceResourceId;
            var ServiceEndpointUri = dcr.ServiceEndpointUri;
            SaveInCache("LastLoggedInUser", dcr.UserId);

            SharePointClient client = new SharePointClient(ServiceEndpointUri, async () =>
            {

                Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential creds =
                new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential(
                    disco.AppIdentity.ClientId, disco.AppIdentity.ClientSecret);

                string accessToken = (await disco.AuthenticationContext.AcquireTokenSilentAsync(
                    ServiceResourceId,
                    creds,
                    new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId))).AccessToken;

                return accessToken;

            });

            disco.ClearCache();
            return client;
        }
        public async Task<bool> CreateProject(string projectTitle, string projectOwner, string documentLink, string documentName)
        {
            //code here
        }
        public async Task<string> GetUpdateAccessToken()
        {
            DiscoveryContext disco = GetFromCache("DiscoveryContext") as DiscoveryContext;

            if (disco == null)
            {
                disco = await DiscoveryContext.CreateAsync();
                SaveInCache("DiscoveryContext", disco);
            }

            var dcr = await disco.DiscoverResourceAsync(this.SiteUrl);

            string clientId = disco.AppIdentity.ClientId;
            string clientSecret = disco.AppIdentity.ClientSecret;
            SaveInCache("LastLoggedInUser", dcr.UserId);

            ClientCredential creds = new ClientCredential(clientId, clientSecret);
            UserIdentifier userId = new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId);
            AuthenticationContext authContext = disco.AuthenticationContext;

            AuthenticationResult authResult = await authContext.AcquireTokenSilentAsync(this.SiteUrl, creds, userId);

            return authResult.AccessToken;

        }
        public void SaveInCache(string name, object value)
        {
            System.Web.HttpContext.Current.Session[name] = value;
        }
        public object GetFromCache(string name)
        {
            return System.Web.HttpContext.Current.Session[name];
        }
        public void RemoveFromCache(string name)
        {
            System.Web.HttpContext.Current.Session.Remove(name);
        }

    }

}