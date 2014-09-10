Demos for Module 03 - Hooking into Apps for SharePoint
======================================================

This script is used to demonstrate some sample applications provided in this demo folder. They are not the same solutions built in module 3, but they demonstrate the same things the students will build.

## Demo #1: Programming OAuth

1. Open the solution **WelcomeAppPart\WelcomeAppPart.sln**.
1. Update the **Site URL** to refer to your environment.
1. Rebuild the solution.
1. Start Fiddler.
1. Hit **F5** to run the app.
1. When the app starts, navigate to the host web.
1. Add the app part to the page.
1. Look at the OAuth calls in Fiddler.
1. Place the app part in edit mode.
1. Uncheck the property to show the user's image.
1. Stop debugging.
1. Open **WelcomeUser.aspx.cs** to show the code.
1. Go over the code to show how the call is made and OAuth tokens are used.

## Demo #2: The Cross-Domain Library

1. Open the solution **CrossDomainCRUD\CrossDomainCRUD.sln**.
1. Update the **Site URL** to refer to your environment.
1. Rebuild the solution.
1. Hit **F5** to run the app.
1. Review the list items.
1. Add a new item to the list.
1. Stop debugging.
1. Open **app.js**.
  1. Show how the required libraries are loaded.
  1. Show how the Chrome control is used.
1. Open **csom.listitems.com**.
  1. Show how new items are created using the cross-domain library.