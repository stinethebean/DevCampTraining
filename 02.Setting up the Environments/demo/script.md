# Demos

## Demo1: Obtain an O365 Subscription
Perform this demo online from [Office Dev Center](http://msdn.microsoft.com/en-us/library/office/fp179924(v=office.15).aspx)

1. Navigate to the [Office Dev Center](http://msdn.microsoft.com/en-us/library/office/fp179924(v=office.15).aspx)
2. Under the heading **Sign up for an Office 365 Developer Site** click **Try It Free**.
3. Fill out the form to obtain your trial O365 subscription.
4. When completed, you will have a developer site in the [subscription].sharepoint.com domain located at the root of your subscription (e.g. https://mysubscription.sharepoint.com)

## Demo2: Obtain an Azure subscription
Perform this demo online from [Azure Portal](https://manage.windowsazure.com) using the credentials from Demo 1

1. Navigate to the [Azure Portal](https://manage.windowsazure.com)
2. If prompted, log in using the credentials you created for your O365 subscription.
3. After logging in, you should see a screen notifying you that you do not have a subscription
4. Click Sign Up for Windows Azure.
5. Fill out the form to obtain your free trial.

## Demo3: Create a Provider-Hosted App

1. Launch **Visual Studio 2013** as administrator. 
2. In Visual Studio select **File/New/Project**.
3. In the New Project dialog:
  1. Select **Templates/Visual C#/Office/SharePoint/Apps**.
  2. Click **App for SharePoint 2013**.
  3. Name the new project **AzureCloudApp** and click **OK**.
4. In the New App for SharePoint wizard:
  1. Enter the address of a SharePoint site to use for testing the app (***NOTE:*** The targeted site must be based on a Developer Site template)
  2. Select **Provider-Hosted** as the hosting model.
  3. Click **Next**.
  4. Select **ASP.NET MVC Web Application**.
  5. Click **Next**.
  6. Select the option labeled **Use Windows Azure Access Control Service (for SharePoint cloud apps)**.
  7. Click **Finish**.
  8. When prompted, log in using your O365 administrator credentials.
5. Press **F5** to show it works.
6. Right-click the remote web project and select **Publish**.
  1. Click **Windows Azure Web Sites**
  2. Click **New**
  3. Fill out information for new site
  4. Click **Create**
7. Log into the O365 developer site as an administrator
  1. From the developer site, navigate to **/_layouts/15/appregnew.aspx**.
  2. Click **Generate** next to Client ID.
  3. Click **Generate** next to Client Secret.
  4. Enter **Azure Cloud App** as the Title.
  5. Enter the **App Domain** for the Azure web site you created earlier (e.g., azurecloudapp.azurewebsites.net)
  6. Enter the **Redirect URI** as the reference for the Customers page (e.g. https://azurecloudapp.azurewebsites.net/Customers).
  7. Click **Create.**
  8. Save the **Client ID** and **Client Secret** separately for later use.
8. In the app project open the **AppManifest.xml** file in a text editor.
9. Update the **Client ID** and **App Start page**.
10. Open the **web.config** file for the **AzureCloudAppWeb** project.
11. Update the **Client ID** and **Client Secret** to use the generated values.
12. Log into Azure and create the AppSettings
  1. Return to the [Azure Management portal](https://manage.windowsazure.com).
  2. Click **Web Sites**.
  3. Select your Azure Web Site.
  4. Click **Configure**.
  5. In the **App Settings** section, add a **ClientId** and **ClientSecret** setting.
  6. Set the values to the values you generated earlier.
  7. Click **Save**.

