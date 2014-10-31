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
3. Click **OK**.<br/>
  ![](img/24.png?raw=true "Figure 24") 

Now you have connected services available to your web application.

###Task 3 - Add Pre-coded Files to the Project
Follow these steps to some required supporting files that are already coded for you.

1. In the **Solution Explorer**, right click the **Models** folder and select **Add/Existing Item**.
2. Navigate to the folder **src\Lab Files\Models**.
3. Select all of the files in the folder and click **Add**.

Now you have some pre-coded supporting files in the project.

###Task 4 - Code the Repository Class
Follow these steps to fill in some missing code in the project for reading and writing using the Office 365 APIs.

1. In the **Models** folder, locate and open the **ProjectRepository.cs** file.
2. At the top of the file, update the **SiteUrl** member with your tenancy information.
3. Add the following code to the **GetMyContacts** method to retrieve contacts from Exchange.

  ```C#
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

  ```
4. Add the following code to the **GetMyFiles** method to retrieve documents from OneDrive for Business.

  ```C#
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
  ```

5. Add the following code to the **CreateProject** method to add a new project to the Research Projects list.

  ```C#
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
  ```

Now the Repository Class is coded to use the Office 365 APIs.

###Task 5 - Create a View for adding Projects
Follow these steps to create a new view that will support adding projects.

1. In the **Solution Explorer**, expand the **Controllers** folder and open **ProjectController.cs**.
2. Right click inside the **Index** method and select **Add View**.
3. In the **Add View** dialog:
  1. Select **Create** for the **Template**.
  2. Select **ViewModel** as the **Model Class**.
  3. Click **Add**.<br/>
    ![](img/25.png?raw=true "Figure 25") 
4. **Replace** all of the code in the view with the following:

  ```HTML
  @model MVCResearchTracker.Models.ViewModel

  @{
      ViewBag.Title = "Add Project";
  }
  
  <h2>Add Project</h2>
  
  
  @using (Html.BeginForm()) 
  {
    @Html.AntiForgeryToken()
    
    <div class="form-horizontal">
        <h4>Add Project</h4>
        <hr />
        @Html.ValidationSummary(true, "", new { @class = "text-danger" })
        <div class="form-group">
            @Html.LabelFor(model => model.Project.Title, htmlAttributes: new { @class = "control-label col-md-2" })
            <div class="col-md-10">
                @Html.EditorFor(model => model.Project.Title, new { htmlAttributes = new { @class = "form-control" } })
                @Html.ValidationMessageFor(model => model.Project.Title, "", new { @class = "text-danger" })
            </div>
        </div>

        <div class="form-group">
            @Html.LabelFor(m => m.Project.Owner, htmlAttributes: new { @class = "control-label col-md-2" })
            <div class="col-md-10">
                @Html.DropDownListFor(m => m.Project.Owner, new SelectList(Model.Contacts, "DisplayName", "DisplayName"), new { htmlAttributes = new { @class = "form-control" } })
                @Html.ValidationMessageFor(m => m.Project.Owner, "", new { @class = "text-danger" })
            </div>
        </div>

        <div class="form-group">
            @Html.LabelFor(m => m.Project.DocumentName, htmlAttributes: new { @class = "control-label col-md-2" })
            <div class="col-md-10">
                @Html.DropDownListFor(m => m.Project.DocumentLink, new SelectList(Model.Files, "Url", "Name"), new { htmlAttributes = new { @class = "form-control" } })
                @Html.ValidationMessageFor(m => m.Project.DocumentName, "", new { @class = "text-danger" })
            </div>
        </div>

        <div class="form-group">
            <div class="col-md-offset-2 col-md-10">
                <input type="submit" value="Create" class="btn btn-default" />
            </div>
        </div>
    </div>
  }

  <div>
      @Html.ActionLink("Back to List", "Index")
  </div>

  @section Scripts {
      @Scripts.Render("~/bundles/jqueryval")
  }

  ```

Now you have finished creating a view for adding projects

###Task 6 - Test the Project
Follow these steps to test your project.

1. Press **F5** to begin debugging.
2. When prompted, log in with your **Organizational Account**.
3. When the application starts, manually navigate to **/Project** by typing in the address bar.
4. In the **Add Project** form:
  1. Name the new project **My New Project**.
  2. Select an **Owner**.
  3. Select an associated **Statement Document**.
  4. Click **Create**.<br/>
    ![](img/26.png?raw=true "Figure 26") 
5. Open a new browser window and navigate to your **Research projects** list in SharePoint.
6. Verify that the new project was successfully added.
  ![](img/27.png?raw=true "Figure 27") 

By completing this lab, you learnt to
- Use Office 365 APIs in a web application
- Register web applications in Azure Active Directory
- Grant permissions to an application