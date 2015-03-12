package com.microsoft.o365_tasks;

import com.microsoft.o365_tasks.auth.AuthManager;
import com.microsoft.o365_tasks.sharepoint.ExtendedListClient;
import com.microsoft.sharepointservices.Credentials;
import com.microsoft.sharepointservices.ListClient;
import com.microsoft.sharepointservices.http.OAuthCredentials;

import android.app.Application;
import android.util.Log;

public class TasksApplication extends Application {
    
    private static final String TAG = "TodoApplication";
    
    private AuthManager mAuthManager;
    
    public AuthManager getAuthManager() {
        
        if (mAuthManager == null) {
            try {
                mAuthManager = new AuthManager(this);
            }
            catch (Exception e) {
                Log.e(TAG, "Error creating authentication context", e);
                throw new RuntimeException("Error creating authentication context", e);
            }
        }
        
        return mAuthManager;
    }


    public ExtendedListClient createExtendedListClient() {

        String accessToken = getAuthManager().getAccessToken();

        Credentials credentials = new OAuthCredentials(accessToken);

        return new ExtendedListClient(Constants.SHAREPOINT_URL, Constants.SHAREPOINT_SITE_PATH, credentials);
    }

    public ListClient createListClient() {

        String accessToken = getAuthManager().getAccessToken();

        Credentials credentials = new OAuthCredentials(accessToken);

        return new ListClient(Constants.SHAREPOINT_URL, Constants.SHAREPOINT_SITE_PATH, credentials);
    }

}
