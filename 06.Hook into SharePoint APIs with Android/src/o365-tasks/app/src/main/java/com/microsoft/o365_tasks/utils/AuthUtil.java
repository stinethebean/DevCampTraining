package com.microsoft.o365_tasks.utils;

import android.app.Activity;
import android.util.Log;

import com.microsoft.o365_tasks.BuildConfig;
import com.microsoft.o365_tasks.TasksApplication;
import com.microsoft.o365_tasks.auth.AuthCallback;
import com.microsoft.o365_tasks.auth.AuthManager;


public class AuthUtil {

    protected static final String TAG = "ActivityUtil";
    
    public static void ensureAuthenticated(final Activity activity, final AuthHandler handler) {

        final TasksApplication application = (TasksApplication) activity.getApplication();
        final AuthManager authManager = application.getAuthManager();
        
        authManager.refresh(new AuthCallback() {
            @Override
            public void onFailure(String errorDescription) {
                handler.onFailure(errorDescription);
            }

            @Override
            public void onCancelled() {
                //Not used
                Log.w(TAG, "AuthManager.refresh failed with onCancelled");
                if (BuildConfig.DEBUG) {
                    throw new RuntimeException("Invalid operation");
                }
            }

            @Override
            public void onSuccess() {
                handler.onSuccess();
            }
        });
    }

    public static interface AuthHandler {
    
        void onFailure(String errorDescription);
    
        void onSuccess();
        
    }   
}
