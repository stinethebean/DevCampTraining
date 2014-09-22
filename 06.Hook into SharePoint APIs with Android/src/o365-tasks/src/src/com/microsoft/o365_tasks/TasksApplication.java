package com.microsoft.o365_tasks;

import com.microsoft.o365_tasks.auth.AuthManager;
import com.microsoft.o365_tasks.sharepoint.SharepointListsClient2;
import com.microsoft.office365.http.OAuthCredentials;

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

    public SharepointListsClient2 createListsClient() {
        
        OAuthCredentials credentials = getAuthManager().getOAuthCredentials();

        return new SharepointListsClient2(Constants.SHAREPOINT_URL, Constants.SHAREPOINT_SITE_PATH, credentials);
    }

}
