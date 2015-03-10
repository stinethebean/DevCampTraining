package com.microsoft.o365_tasks.utils.auth;

import com.microsoft.o365_tasks.R;
import com.microsoft.o365_tasks.utils.IntentUtil;
import com.microsoft.o365_tasks.utils.AuthUtil.AuthHandler;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.content.DialogInterface.OnClickListener;

public abstract class DefaultAuthHandler implements AuthHandler {
    
    private Activity mActivity;

    public DefaultAuthHandler(Activity activity) {
        mActivity = activity;
    }
    
    public void onFailure(String errorDescription) {
        
        //Authentication has failed! Launch a dialog to let the user know.
        //When the user taps Continue, we will restart the app so that they may authenticate again.
        
        new AlertDialog.Builder(mActivity)
            .setTitle(R.string.dialog_session_expired)
            .setMessage(R.string.dialog_auth_failed_message)
            .setPositiveButton(R.string.label_continue, new OnClickListener() {
                @Override public void onClick(DialogInterface dialog, int which) {
                    
                    //user has confirmed - restart the app
                    mActivity.startActivity(IntentUtil.makeRestartForAuthIntent(mActivity));
                }
            })
            .setCancelable(false)
            .create()
            .show();        
    }
}