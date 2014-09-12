Module 03 - Hook into Apps for SharePoint
=========================================

##Overview
In this lab, you will create apps that use both OAuth security and the cross-domain library. You will examine the security flow to better understand the available options.

##Objectives
- Understand the OAuth flow in a Provider-Hosted app
- Understand how to use the Cross-Domain Library in a Provider-Hosted app

##Prerequisites
- Visual Studio 2013 for Windows 8
- You must have an Office 365 tenant and Windows Azure subscription to complete this lab.
- You must have [Fiddler] (http://www.telerik.com/fiddler) installed.

##Exercises
The hands-on lab includes the following exercises:<br/>
1. <a href="#Exercise1">OAuth in a Provider-Hosted App</a><br/>
2. <a href="#Exercise2">Cross-Domain Library in a Provider-Hosted App</a><br/>

<a name="Exercise1"></a>
##Exercise 1: OAuth in a Provider-Hosted App
In this exercise you create a new provider-hosted app and examine the OAuth flow.

###Task 1 - Create the new solution in Visual Studio 2013
Follow these steps to create a new project.

1. Launch **Visual Studio 2013** as administrator.
2. In Visual Studio select **File/New/Project**.
3. In the New Project dialog:
  1. Select **Templates/Visual C#/Office/SharePoint/Apps**.
  2. Click **App for SharePoint 2013**.
  3. Name the new project **ProviderHostedOAuth** and click **OK**.<br/>
     ![](img/01.png?raw=true "Figure 1")
4. In the New App for SharePoint wizard:
  1. Enter the address of a SharePoint site to use for testing the app (***NOTE:*** The targeted site must be based on a Developer Site template)
  2. Select **Provider-Hosted** as the hosting model.
  3. Click **Next**.<br/>
     ![](img/02.png?raw=true "Figure 2")
  4. Select **ASP.NET MVC Web Application**.
  5. Click **Next**.<br/>
     ![](img/03.png?raw=true "Figure 3")
  6. Select the option labeled **Use Windows Azure Access Control Service (for SharePoint cloud apps)**.
  7. Click **Finish**.<br/>
     ![](img/04.png?raw=true "Figure 4")
  8. When prompted, log in using your O365 administrator credentials.
  9. After the new project is created, set breakpoints in **HomeController.cs** as shown.<br/>
    ![](img/05.png?raw=true "Figure 5")

Now you have a new app created.

###Task 2 - Examine the OAuth flow
Follow these steps to examine the OAuth flow of your app while debugging.

1. Start **Fiddler** to capture web traffic from your app.
  1. In Fiddler click **Tools/Fiddler Options**.
  2. Click **HTTPS**.
  3. Check the box entitled **Decrypt HTTPS Traffic**.
  4. When warned, click **Yes** to trust the Fiddler root certificate.
  5. Confirm any additional dialog boxes to install the certificate.
  6. Click **OK** to close the options dialog.
2. Debug the app by pressing **F5** in Visual Studio 2013.
  1. When prompted, sign into Office 365.
  2. When prompted, click **Trust It**.<br/>
      ![](img/07.png?raw=true "Figure 7")
  3. When the first breakpoint is hit, look for the session in Fiddler near the bottom of the list.<br/>
      ![](img/08.png?raw=true "Figure 8")
  4. Right click the session and select **View in New Window**.
  5. Click the **Web Forms** tab.
  6. Notice that SharePoint has included the SPHostUrl, SPLanguage, SPClientTag, and SPProductNumber query string parameters in the initial call. These are known as the **Standard Tokens**.
  7. Notice that the context token is included in the body as **SPAppToken**<br/>.
      ![](img/09.png?raw=true "Figure 9")
  8. Close the window.
  9. Return to Visual Studio, and press **F5** to continue debugging.
  10. When the second breakpoint is hit, look for the session in Fiddler near the bottom of the list.<br/>
      ![](img/10.png?raw=true "Figure 10")
  11. Right click the session and select **View in New Window**.
  12. Click the **Headers** tab and examine the access token in the **Cookies/Login** section <br/>
      ![](img/11.png?raw=true "Figure 11")
  13. Return to Visual Studio, and press **F5** to continue debugging.
  14. With the app still running, open a new browser window to **/_layouts/15/AppPrincipals.aspx**.
  15. Look for **ProviderHostedOAuth** in the list of registered apps to confirm that the app was registered during debugging.
  16. Stop debugging.

Now you have debugged the app and followed the OAuth flow.

<a name="Exercise2"></a>
##Exercise 2: Cross-Domain Library in a Provider-Hosted App
In this exercise you use the cross-domain library to access a list in the app web.

###Task 1 - Create the new ASP.NET Web Forms app project in Visual Studio 2013
Follow these steps to create a new project.

1. Launch **Visual Studio 2013** as administrator.
2. In Visual Studio select **File/New/Project**.
3. In the New Project dialog:
  1. Select **Templates/Visual C#/Office/SharePoint/Apps**.
  2. Click **App for SharePoint 2013**.
  3. Name the new project **DeepDiveCloudApp** and click **OK**.<br/>
     ![](img/12.png?raw=true "Figure 1")
4. In the New App for SharePoint wizard:
  1. Enter the address of a SharePoint site to use for testing the app (***NOTE:*** The targeted site must be based on a Developer Site template)
  2. Select **Provider-Hosted** as the hosting model.
  3. Click **Next**.<br/>
     ![](img/13.png?raw=true "Figure 2")
  4. Select **ASP.NET Web Forms Application**.
  5. Click **Next**.<br/>
     ![](img/14.png?raw=true "Figure 3")
  6. Select the option labeled **Use Windows Azure Access Control Service (for SharePoint cloud apps)**.
  7. Click **Finish**.<br/>
     ![](img/15.png?raw=true "Figure 4")
  8. When prompted, log in using your O365 administrator credentials.
  9. Open **Default.aspx.cs** from the **DeepDiveCloudAppWeb** project.
  10. **Delete** the code that is used to obtain the host web title so your code appears as follows:<br/>
     ![](img/16.png?raw=true "Figure 5")

Now you have created a new ASP.NET Web Forms app project.

## Task 2 -  Using the Cross-Domain Library
Follow these steps to access a list in the app web.

1. Right click the **DeepDiveCloudApp** project and select **Add/New Item**.
2. In the **Add New Item** dialog, select **List**.
3. name the new list **Terms**.
4. Click **Add**.<br/>
       ![](img/17.png?raw=true "Figure 7")
5. In the **SharePoint Customization Wizard**, select **Create list instance based on existing list template**.
6. Select **Custom List**.
7. Click **Finish**.<br/>
       ![](img/18.png?raw=true "Figure 8")
8. Open the **Elements.xml** file associated with the new list instance **DeepDiveCloudApp/Terms/Elements.xml**.
9. Add the following XML inside the **ListInstance** element to pre-populate the list with data.

    ````xml
    <Data>
      <Rows>
        <Row>
          <Field Name="Title">SharePoint-Hosted App</Field>
        </Row>
        <Row>
          <Field Name="Title">Provider-Hosted App</Field>
        </Row>
        <Row>
          <Field Name="Title">Microsoft Azure</Field>
        </Row>
        <Row>
          <Field Name="Title">Office 365</Field>
        </Row>
        <Row>
          <Field Name="Title">SharePoint Online</Field>
        </Row>
      </Rows>
    </Data>
    ````

10. Right click the **Scripts** folder in the **DeepDiveCloudAppWeb** project and select **Add/New/JavaScript File**.
11. Name the new file **crossdomain**.
12. Click **OK**.
13. **Add** the following code to **crossdomain.js** to read the Terms list items.

    ````javascript
    (function () {
        "use strict";

        jQuery(function () {

            //Get Host and App web URLS
            var appWebUrl = "";
            var spHostUrl = "";
            var args = window.location.search.substring(1).split("&");

            for (var i = 0; i < args.length; i++) {
                var n = args[i].split("=");
                if (n[0] == "SPHostUrl")
                    spHostUrl = decodeURIComponent(n[1]);
            }

            for (var i = 0; i < args.length; i++) {
                var n = args[i].split("=");
                if (n[0] == "SPAppWebUrl")
                    appWebUrl = decodeURIComponent(n[1]);
            }

            //Load Libraries
            var scriptbase = spHostUrl + "/_layouts/15/";

            jQuery.getScript(scriptbase + "SP.RequestExecutor.js", function (data) {

                //Call Host Web with REST
                var executor = new SP.RequestExecutor(appWebUrl);
                executor.executeAsync({
                    url: appWebUrl + "/_api/web/lists/getbytitle('Terms')/items",
                    method: "GET",
                    headers: { "accept": "application/json;odata=verbose" },
                    success: function (data) {

                        var results = JSON.parse(data.body).d.results;
                        for (var i = 0; i < results.length; i++) {
                            $("#termList").append("<li>" + results[i].Title + "</li>");
                        }
                    },
                    error: function () {
                        alert("Error!");
                    }
                });

            });

        });

    }());
    ````

14. Open **Default.aspx** from the **DeepDiveCloudAppWeb** project.
15. Add the following script references in the **head** section.

    ````xml
    <script src="../Scripts/jquery-1.9.1.js"></script>
    <script src="../Scripts/crossdomain.js"></script>
    ````

16. Add an unordered list element to display the terms. The list element is shown in context below.

    ````xml
    <body>
    <form id="form1" runat="server">
        <div id="chrome_ctrl_placeholder"></div>
        <div>
            <ul id="termList"></ul>
        </div>
    </form>
    </body>
    ````

17. Press **F5** to debug the app.<br/>
     ![](img/19.png?raw=true "Figure 9")

Now you have read list items using the Cross-Domain Library.

##Summary
By completing this hands-on lab you learnt how to:
- Program OAuth in a Provider-Hosted app.
- Utilize the Cross-Domain library in a Provider-Hosted app.
