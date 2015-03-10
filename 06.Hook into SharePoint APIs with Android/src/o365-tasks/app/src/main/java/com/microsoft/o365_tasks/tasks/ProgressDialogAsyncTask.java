package com.microsoft.o365_tasks.tasks;

import android.app.ProgressDialog;
import android.content.Context;
import android.os.AsyncTask;

/**
 * A simple AsyncTask class which launches an "in progress" dialog while the background task is running.
 *
 * @param <T> the return type of doInBackground
 */
public abstract class ProgressDialogAsyncTask<T> extends AsyncTask<Void, Void, T> {

    private ProgressDialog mDialog;

    protected ProgressDialogAsyncTask(Context context) {
        mDialog = new ProgressDialog(context);
    }
    
    protected ProgressDialogAsyncTask(Context context, CharSequence message) {
        this(context);
        setDialogMessage(message);
    }
    
    protected ProgressDialogAsyncTask(Context context, CharSequence message, CharSequence title) {
        this(context, message);
        setDialogTitle(title);
    }
    
    public void setDialogTitle(CharSequence title) {
        mDialog.setTitle(title);
    }
    
    public void setDialogMessage(CharSequence message) {
        mDialog.setMessage(message);
    }
    
    @Override
    protected void onPreExecute() {
        //Start the dialog
        mDialog.setCancelable(false);
        mDialog.setIndeterminate(true);
        mDialog.show();
    }

    @Override
    protected void onPostExecute(T result) {
        onResult(result);
        //Dismiss the dialog
        if (mDialog.isShowing()) {
            mDialog.dismiss();
        }
    }

    protected abstract void onResult(T result);
}
