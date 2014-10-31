Module 04 - Hook into Office 365 APIs
=====================================

##Overview
In this lab, you will create a web application that uses the Office 365 APIs. The lab will create a "Research Tracker" that allows you to define new research projects in a SharePoint list, assign an owner, and create a project statement.

##Objectives
- Learn to use Office 365 APIs in a web application
- Understand how to register web applications in Azure Active Directory
- Understand how to grant permissions to an application

##Prerequisites
- Visual Studio 2013 for Windows 8
- You must have an Office 365 tenant and Microsoft Azure subscription to complete this lab.
- You must have completed the lab associated with Module 2.

##Exercises
The hands-on lab includes the following exercises:<br/>
1. <a href="#Exercise1">Prepare Data Sources</a><br/>
2. <a href="#Exercise2">Create an MVC5 Web Application</a><br/>

<a name="Exercise1"></a>
##Exercise 1: Prepare Data Sources
In this exercise you will add data and documents to Office 365 for use in your solution.

###Task 1 - Create a Research Projects List
Follow these steps to create a Research Projects List to project definitions.

1. Log into SharePoint Online using your **Organizational Account**.
2. Click **Site Contents**.
3. Click **Add an App**.
4. Click **Custom List**.
5. Name the new list **Research Projects**.
6. Click **Create**.<br/>
  ![](img/01.png?raw=true "Figure 1")
7. Click on the newly-create **Research Projects** list.
8. Click the **List** tab and then **List Settings**.<br/>
  ![](img/02.png?raw=true "Figure 2")
9. Click **Create Column**.<br/>
  ![](img/03.png?raw=true "Figure 3")
10. Create a new column as a **Single Line of Text** named **Owner**.<br/>
  ![](img/04.png?raw=true "Figure 4")
11. Click **OK**.
12. Click **Create Column**.
13. Create a new column as a **Hyperlink or Picture** named **Statement**.<br/>
  ![](img/05.png?raw=true "Figure 5")
14. Click **OK**.

Now you have created the required list for the project.

###Task 2 - Upload a Project Statement
Follow these steps to upload a project statement to your OneDrive for Business library.

1. Log into SharePoint Online using your **Organizational Account**.
2. Click **OneDrive**<br/>
  ![](img/06.png?raw=true "Figure 6")
3. If you are presented with the welcome screen, simply click **Next**.<br/>
  ![](img/07.png?raw=true "Figure 7")
4. Click **Upload**.<br/>
  ![](img/09.png?raw=true "Figure 9")
5. Locate the **ProjectStatement.docx** file, which can be found in the **src/Lab Files/Docs** folder.
6. Select the file and click **Open**.<br/>
  ![](img/10.png?raw=true "Figure 10")

Now you have a document available in OneDrive for Business.

###Task 2 - Create some contacts
Follow these steps to create some contacts in Outlook.

1. Log into SharePoint Online using your **Organizational Account**.
2. Click **Outlook**.<br/>
  ![](img/11.png?raw=true "Figure 11")
3. If prompted, select your settings and click **Save**.<br/>
  ![](img/12.png?raw=true "Figure 12")
4. Click **People**.<br/>
  ![](img/13.png?raw=true "Figure 13")
5. Click **New**.<br/>
  ![](img/14.png?raw=true "Figure 14")
6. Click **Create Contact**.<br/>
  ![](img/15.png?raw=true "Figure 15")
7. Create a new contact being sure to fill out **First Name**, **Last Name** and **Email**.
8. Click **Save**.<br/>
  ![](img/16.png?raw=true "Figure 16") 
9. You may optionally add a few more contacts, but only one is necessary for the lab.

Now you have a contact to work with in the project.

<a name="Exercise2"></a>
##Exercise 2: Create an MVC5 Web Application
In this exercise you will create a web application that uses the Office 365 APIs to interact with the data sources you created earlier.

###Task 1 - Create a new Project
Follow these steps to create a new MVC5 web application project.

1. Start **Visual Studio 2013**.
2. Select **File/New/Project** from the main menu.
3. In the **New Project** dialog:
  1. Select **C#/Web**.
  2. Click **ASP.NET Web Application**.
  3. Name the new project **MVCResearchTracker**.
  4. Click **OK**.<br/>
    ![](img/17.png?raw=true "Figure 17") 
4. In the **New ASP.NET Project** dialog:
  1. Click the **MVC** template.
  2. Click **Change Authentication**.
  3. Click **Organizational Accounts**.
  4. Enter the domain for your Office 365 tenancy.
  5. Click **OK**.<br/>
    ![](img/18.png?raw=true "Figure 18") 
  6. When prompted, sign into Azure Active Directory with your **Organizational Account**.
  7. Click **OK**.<br/>
    ![](img/19.png?raw=true "Figure 19")
5. Press **F5** to test your application security by logging in with your **Organizational Account**.
 
Now you have an ASP.NET MVC web application secured with an Organizational Account.

###Task 2 - Install the Office 365 API Tools
Follow these steps to install the Office 365 API Tools, if you do not have them already.

1. In Visual Studio 2013, select **Tools/Extensions and Updates** from the main menu.
2. In the **Extensions and Updates** dialog, enter **Office 365** in the **Search** box.
3. Locate the **Office 365 API Tools** and click **Install**.
4. Click **Close**.<br/>
  ![](img/20.png?raw=true "Figure 20") 

Now you have the Office 365 API tools installed.

###Task 3 - Add Connected Services
Follow these steps to add connected services to the MVC5 web application project.

1. In the **Solution Explorer**, right click the **MVCResearchTracker** project and select **Add/Connected Service**.
2. In the **Services Manager** dialog:
  1. Click **Register Your App**.
  2. When prompted, sign into Azure Active Directory with your **Organizational Account**.
  3. Click **Contacts**.
  4. Click **Permissions**.
  5. Check **Read Users' Contacts**.
  6. Click **Apply**.<br/>
    ![](img/21.png?raw=true "Figure 21") 
  7. Click **My Files**.
  8. Click **Permissions**.
  9. Check **Read Users' Files**.
  10. Click **Apply**.<br/>
    ![](img/22.png?raw=true "Figure 22") 
  11. Click **Sites**.
  12. Click **Permissions**.
  13. Check **Create or Delete Items and Lists in All Site Collections**.
  14. Check **Edit or Delete Items in All Site Collections**.
  15. Check **Read Items in All Site Collections**.
  16. Click **Apply**.<br/>
    ![](img/23.png?raw=true "Figure 23")
4. Click **App Properties**.
5. **Delete** the Redirect Uri that uses *http** leaving in place the one that uses **https**.
6. Click **Apply**.<br/>
  ![](img/24a.png?raw=true "Figure 24a") 
7. Click **OK**.<br/>
  ![](img/24.png?raw=true "Figure 24") 

Now you have connected services available to your web application.

###Task 3 - Add Pre-coded Files to the Project
Follow these steps to some required supporting files that are already coded for you.

1. In the **Solution Explorer**, right click the **Models** folder and select **Add/Existing Item**.
2. Navigate to the folder **src\Lab Files\Models**.
3. Select all of the files in the folder and click **Add**.
4. In the **Solution Explorer**, right click the **Utils** folder and select **Add/Existing Item**.
5. Navigate to the folder **src\Lab Files\Utils**.
6. Select the **Helper.cs** file in the folder and click **Add**.

Now you have some pre-coded supporting files in the project.

###Task 4 - Stub Out the Home Controller code
Follow these steps to add methods to the Home Controller that you will fill out throughout the lab.

1. In the **Solution Explorer**, expand the **Controllers** folder, and open **HomeController.cs** for editing.
2. **Replace** all of the code with the following:

  ```C#
  using System;
  using System.Collections.Generic;
  using System.Linq;
  using System.Web;
  using System.Web.Mvc;
  using Microsoft.IdentityModel.Clients.ActiveDirectory;
  using Microsoft.Office365.Discovery;
  using Microsoft.Office365.OutlookServices;
  using System.Configuration;
  using System.Threading.Tasks;
  using MVCResearchTracker.Utils;
  using MVCResearchTracker.Models;
  using Microsoft.Office365.SharePoint.CoreServices;
  using System.Text;
  using System.Xml.Linq;
  using System.Net.Http;
  using System.Net.Http.Headers;

  namespace MVCResearchTracker.Controllers
  {
      public class HomeController : Controller
      {
          private const string spSite = "https://[tenancy].sharepoint.com";
          private const string discoResource = "https://api.office.com/discovery/";
          private const string discoEndpoint = "https://api.office.com/discovery/v1.0/me/";

          public async Task<ActionResult> Index(string code)
          {
              return View();
          }

          public async Task<ActionResult> Contacts(string code)
          {
              return View();
          }

          public async Task<ActionResult> Files(string code)
          {
              return View();
          }

          public async Task<ActionResult> Projects(ViewModel submitModel, string code)
          {
              return View();
          }

          public ActionResult Finished()
          {
              return View();
          }
      }
  }

  ```
3. **Replace** the placeholder **[tenancy]** with the name of your SharePoint tenancy.

Now you have code placeholders in the Home Controller.

###Task 5 - Code the Discovery Service
Follow these steps to use the Discovery Service to locate endpoints for Exchnage contacts and OneDrive for Business files.

1. **Replace** the **Index** method in the **HomeController** with the following code:
  ```C#
        public async Task<ActionResult> Index(string code)
        {
            AuthenticationContext authContext = new AuthenticationContext(
               ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
               true);

            ClientCredential creds = new ClientCredential(
                ConfigurationManager.AppSettings["ida:ClientID"],
                ConfigurationManager.AppSettings["ida:Password"]);

            DiscoveryClient disco = Helpers.GetFromCache("DiscoveryClient") as DiscoveryClient;

            //Redirect to login page if we do not have an 
            //authorization code for the Discovery service
            if (disco == null && code == null)
            {
                Uri redirectUri = authContext.GetAuthorizationRequestURL(
                    discoResource,
                    creds.ClientId,
                    new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
                    UserIdentifier.AnyUser,
                    string.Empty);

                return Redirect(redirectUri.ToString());
            }

            //Create a DiscoveryClient using the authorization code
            if (disco == null && code != null)
            {

                disco = new DiscoveryClient(new Uri(discoEndpoint), async () =>
                {

                    var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
                        code,
                        new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
                        creds);

                    return authResult.AccessToken;
                });

            }

            //Discover required capabilities
            CapabilityDiscoveryResult contactsDisco = await disco.DiscoverCapabilityAsync("Contacts");
            CapabilityDiscoveryResult filesDisco = await disco.DiscoverCapabilityAsync("MyFiles");

            Helpers.SaveInCache("ContactsDiscoveryResult", contactsDisco);
            Helpers.SaveInCache("FilesDiscoveryResult", filesDisco);

            List<MyDiscovery> discoveries = new List<MyDiscovery>(){
                new MyDiscovery(){
                    Capability = "Contacts",
                    EndpointUri = contactsDisco.ServiceEndpointUri.OriginalString,
                    ResourceId = contactsDisco.ServiceResourceId,
                    Version = contactsDisco.ServiceApiVersion
                },
                new MyDiscovery(){
                    Capability = "My Files",
                    EndpointUri = filesDisco.ServiceEndpointUri.OriginalString,
                    ResourceId = filesDisco.ServiceResourceId,
                    Version = filesDisco.ServiceApiVersion
                }
            };

            return View(discoveries);

        }
  ```
2. In the **Solution Explorer**, expand the **Views\Home** folder and open **Index.cshtml** for editing.
3. **Replace** the code with the following code:
  ```HTML
  @model IEnumerable<MVCResearchTracker.Models.MyDiscovery>

  @{
      ViewBag.Title = "Discoveries";
  }
  
  <h2>Discoveries</h2>
  
  
  <table class="table">
    <tr>
        <th>
            @Html.DisplayNameFor(model => model.Capability)
        </th>
        <th>
            @Html.DisplayNameFor(model => model.EndpointUri)
        </th>
        <th>
            @Html.DisplayNameFor(model => model.ResourceId)
        </th>
        <th>
            @Html.DisplayNameFor(model => model.Version)
        </th>
        <th></th>
    </tr>

  @foreach (var item in Model) {
    <tr>
        <td>
            @Html.DisplayFor(modelItem => item.Capability)
        </td>
        <td>
            @Html.DisplayFor(modelItem => item.EndpointUri)
        </td>
        <td>
            @Html.DisplayFor(modelItem => item.ResourceId)
        </td>
        <td>
            @Html.DisplayFor(modelItem => item.Version)
        </td>
    </tr>
  }

  </table>

  <div>
    @Html.ActionLink("Get Contacts", "Contacts")
  </div>
  ```
4. Press **F5** to debug the project.
5. When prompted, 

By completing this lab, you learnt to
- Use Office 365 APIs in a web application
- Register web applications in Azure Active Directory
- Grant permissions to an application