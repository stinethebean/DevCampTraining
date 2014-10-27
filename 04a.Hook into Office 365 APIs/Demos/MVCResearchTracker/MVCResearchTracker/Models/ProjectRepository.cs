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
        private string SiteUrl = "https://devcamp2014.sharepoint.com";
        public async Task<List<MyContact>> GetMyContacts()
        {
            var client = await EnsureExchangeClientCreated();

            var contactsResults = await client.Me.Contacts.ExecuteAsync();

            var contactList = new List<MyContact>();

            foreach (var contact in contactsResults.CurrentPage.OrderBy(c => c.Surname))
            {
                contactList.Add(new MyContact
                {
                    Id = contact.Id,
                    GivenName = contact.GivenName,
                    Surname = contact.Surname,
                    DisplayName = contact.Surname + ", " + contact.GivenName,
                    CompanyName = contact.CompanyName,
                    EmailAddress1 = contact.EmailAddress1,
                    BusinessPhone1 = contact.BusinessPhone1,
                    HomePhone1 = contact.HomePhone1
                });
            }
            return contactList;
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
            var client = await EnsureFilesClientCreated();

            var filesResults = await client.Files["Shared with Everyone"].ToFolder().Children.ExecuteAsync();

            var fileList = new List<MyFile>();

            foreach (var file in filesResults.CurrentPage.OrderBy(e => e.Name))
            {
                fileList.Add(new MyFile
                {
                    Id = file.Id,
                    Name = file.Name,
                    Url = file.Url
                });
            }
            return fileList;
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
            StringBuilder requestUri = new StringBuilder()
                 .Append(this.SiteUrl)
                 .Append("/_api/web/lists/getbyTitle('Research Projects')/items");

            string accessToken = GetFromCache("UpdateToken") as string;

            XNamespace atom = "http://www.w3.org/2005/Atom";
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";

            XElement message = new XElement(atom + "entry",
                new XAttribute(XNamespace.Xmlns + "d", d),
                new XAttribute(XNamespace.Xmlns + "m", m),
                new XElement(atom + "category", new XAttribute("term", "SP.Data.Research_x0020_ProjectsListItem"), new XAttribute("scheme", "http://schemas.microsoft.com/ado/2007/08/dataservices/scheme")),
                new XElement(atom + "content", new XAttribute("type", "application/xml"),
                    new XElement(m + "properties",
                        new XElement(d + "Statement", new XAttribute(m + "type", "SP.FieldUrlValue"),
                            new XElement(d + "Description", documentName),
                            new XElement(d + "Url", documentLink)),
                        new XElement(d + "Title", projectTitle),
                        new XElement(d + "Owner", projectOwner))));

            StringContent requestData = new StringContent(message.ToString());

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUri.ToString());
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            requestData.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/atom+xml");
            request.Content = requestData;
            HttpResponseMessage response = await client.SendAsync(request);
            return true;
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