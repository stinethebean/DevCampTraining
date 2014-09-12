Demos from Module 02 - Setting up the Environments
==================================================

This script is used to demonstrate what attendees will do in module 2.

## Demo #1: Obtain an Office 365 Subscription

> Perform this demo online from [Office Dev Center](http://msdn.microsoft.com/en-us/library/office/fp179924.aspx)

1. Navigate to the [Office Dev Center](http://msdn.microsoft.com/en-us/library/office/fp179924.aspx)
1. Under the heading **Sign up for an Office 365 Developer Site** click **Try It Free**.
1. Fill out the form to obtain your trial Office 365 subscription.
1. When completed, you will have a developer site in the **[subscription].sharepoint.com domain** located at the root of your subscription (e.g. https://mysubscription.sharepoint.com).

## Demo #2: Obtain an Azure subscription

> Perform this demo online from [Azure Management Portal](https://manage.windowsazure.com) using the credentials from the [first demo](#demo-1-obtain-an-office-365-subscription).

1. Navigate to the [Azure Management Portal](https://manage.windowsazure.com).
1. If prompted, log in using the credentials you created for your Office 365 subscription.
1. After logging in, you should see a screen notifying you that you do not have a subscription.
1. Click Sign Up for Windows Azure.
1. Fill out the form to obtain your free trial.

## Demo #3: Create a Provider-Hosted App

> This demonstration is used to validate the two previous demos have been performed correctly. It is considered optional for the instructor.

1. Launch **Visual Studio 2013** as administrator.
1. In Visual Studio select **File/New/Project**.
1. In the New Project dialog:
  1. Select **Templates/Visual C#/Office/SharePoint/Apps**.
  1. Click **App for SharePoint 2013**.
  1. Name the new project **AzureCloudApp** and click **OK**.
1. In the New App for SharePoint wizard:
  1. Enter the address of a SharePoint site to use for testing the app (***NOTE:*** The targeted site must be based on a **Developer Site** template)
  1. Select **Provider-Hosted** as the hosting model.
  1. Click **Next**.
  1. Select **ASP.NET MVC Web Application**.
  1. Click **Next**.
  1. Select the option labeled **Use Windows Azure Access Control Service (for SharePoint cloud apps)**.
  1. Click **Finish**.
  1. When prompted, log in using your Office 365 administrator credentials.
1. Press **F5** to show it works.
1. Right-click the remote web project and select **Publish**.
  1. Click **Windows Azure Web Sites**
  1. Click **New**
  1. Fill out information for new site
  1. Click **Create**
1. Log into the Office 365 developer site as an administrator
  1. From the developer site, navigate to **/_layouts/15/appregnew.aspx**.
  1. Click **Generate** next to **Client ID**.
  1. Click **Generate** next to **Client Secret**.
  1. Enter **Azure Cloud App** as the **Title**.
  1. Enter the **App Domain** for the Azure web site you created earlier (e.g., azurecloudapp.azurewebsites.net)
  1. Enter the **Redirect URI** as the reference for the Customers page (e.g. https://azurecloudapp.azurewebsites.net/Customers).
  1. Click **Create.**
  1. Save the **Client ID** and **Client Secret** separately for later use.
1. In the app project open the **AppManifest.xml**.
  1. Update the **Client ID** and **App Start Page**.
1. Open the **web.config** file for the **AzureCloudAppWeb** project.
1. Update the `ClientId` and `ClientSecret` app settings to use the generated values.
1. Log into Azure and create the AppSettings:
  1. Return to the [Azure Management portal](https://manage.windowsazure.com).
  1. Click **Web Sites**.
  1. Select your Azure Web Site.
  1. Click **Configure**.
  1. In the **App Settings** section, add a **ClientId** and **ClientSecret** setting.
  1. Set the values to the values you generated earlier.
  1. Click **Save**.
