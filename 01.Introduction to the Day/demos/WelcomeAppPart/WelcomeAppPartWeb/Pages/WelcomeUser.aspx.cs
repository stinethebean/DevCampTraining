using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Linq;

namespace WelcomeAppPartWeb.Pages
{
    public partial class WelcomeUser : Page
    {
        protected async void Page_Load(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
            string accessToken = spContext.UserAccessTokenForSPHost;

            StringBuilder requestUri = new StringBuilder()
            .Append(Request.QueryString["SPHostUrl"])
            .Append("/_api/web/currentuser");

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUri.ToString());
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            HttpResponseMessage response = await client.SendAsync(request);
            string responseString = await response.Content.ReadAsStringAsync();

            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            XElement root = XElement.Parse(responseString);
            string title = root.Descendants(d + "Title").First().Value;
            WelcomeLabel.Text = title;
        }
    }
}