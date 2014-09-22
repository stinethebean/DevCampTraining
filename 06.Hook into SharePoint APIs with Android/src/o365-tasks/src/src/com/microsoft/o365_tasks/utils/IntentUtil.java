package com.microsoft.o365_tasks.utils;

import android.app.Activity;
import android.content.ComponentName;
import android.content.Intent;

import com.microsoft.o365_tasks.StartActivity;

public class IntentUtil {

    /**
     * Make an Intent that can be used to re-launch this application's task in its base state.
     * 
     * Optionally, the Activity Intent can be configured to next launch the current activity
     * instead of its default behaviour.
     *  
     * @param activity
     * @return
     */
    public static Intent makeRestartForAuthIntent(Activity activity) {
        
        final Intent intent = Intent.makeRestartActivityTask(new ComponentName(activity, StartActivity.class));
        intent.putExtra(StartActivity.PARAM_AUTH_IMMEDIATE, true);
        return intent;
    }
}
