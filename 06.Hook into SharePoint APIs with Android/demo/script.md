# Demos

## Demo1: Authentication with Azure AD
This demo is just a quick look at the code which goes into authenticating with Azure AD.
To prepare for the demo, do the steps in the lab up to the point where you can run the app.

For the demo:
1.  Launch the app and go through the authentication process.

2.  Use the "Clear auth token" function to reset the current access token.

3.  Demo the dialog prompting the user to re-authenticate.

4.  Return to Eclipse and navigate to the com.microsoft.o365_tasks.StartActivity class.

5.  In this class, point out the AuthManager class, and the call to mAuthManager.forceAuthenticate
    when the user clicks "Sign in".

    This is the "app start" part of the pattern described in the slide deck. You may like to
    refer back to this slide.

6.  Navigate to com.microsoft.o365_tasks.auth.AuthManager.forceAuthenticate

    Point out that this method just invokes the acquireToken function from ADAL. The rest of this 
    class deals with the minutia of caching and managing the Access and Refresh tokens returned by
    this function.

7.  Navigate to com.microsoft.o365_tasks.utils.AuthUtil

    This utility function attempts to refresh the current Access token (if necessary). The given
    AuthHandler instance is a callback object which handles the interesting task of alerting the
    user.

8.  Navigate to com.microsoft.o365_tasks.utils.auth.DefaultAuthHandler

    This class is the "default" implementation of the AuthHandler interface. By default
    when the refresh token operation fails this handler will launch a dialog which alerts the
    user and forces them to re-authenticate.

    These two helper classes provide all that we need to implementing the "API calls" part of the
    pattern described in the slide deck. You may like to refer back to this slide.

9.  Navigate to com.microsoft.o365_tasks.ListTasksActivity.optionsActionRefresh

    This function handles the "Refresh" action this Activity.

    The call to ensureAuthenticated accepts a Runnable callback which is invoked only if
    the Access token is valid (or has been refreshed).

    The callback does the actual API call.

10. Navigate to the ensureAuthenticated function.

    This helper function wraps a calls to the two helper functions from earlier.

    The pattern simply requires that every user-initiated action which may result in a call
    to the SharePoint APIs must be preceeded by a call to ensureAuthenticated.



## Demo2: O365 SharePoint for Android

This demo is just a quick look at the code which goes into consuming the SharePoint APIs.
To prepare for the demo, do the steps in the lab up to the point where you can run the app.

For the demo:

01. Navigate to com.microsoft.o365_tasks.data.TaskListItemDataSource

    This class contains all the calls to the SharePoint API.

02. Navigate to com.microsoft.o365_tasks.sharepoint.SharepointListsClient2

    This class extends SharepointListsClient and adds a function for creating lists from a template.
