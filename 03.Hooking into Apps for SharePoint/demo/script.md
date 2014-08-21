# Demos

## Demo1: Programming OAuth

1. Open the solution **WelcomeAppPart.sln**.
2. Update the **Site URL** to refer to your environment.
3. Rebuild the solution.
4. Start Fiddler
5. Hit **F5** to run the app.
6. When the app starts, navigate to the host web.
7. Add the app part to the page
8. Look at the OAuth calls in Fiddler
9. Place the app part in edit mode.
10. Uncheck the property to show the user's image
11. Stop debugging
12. Open WelcomeUser.aspx.cs to show the code
13. Go over the code to show how the call is made and OAuth tokens are used.

## Demo2: The Cross-Domain Library

1. Open the solution **CrossDomainCRUD.sln**.
2. Update the **Site URL** to refer to your environment.
3. Rebuild the solution.
4. Hit **F5** to run the app.
5. Review the list items
6. Add a new item to the list
7. Stop debugging
8. Open **app.js**
  1. Show how the required libraries are loaded
  2. Show how the Chrome control is used
9. Open **csom.listitems.com**
  1. Show how new items are created using the cross-domain library