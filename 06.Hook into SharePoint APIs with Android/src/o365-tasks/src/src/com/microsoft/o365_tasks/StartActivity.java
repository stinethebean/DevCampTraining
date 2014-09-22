package com.microsoft.o365_tasks;

import java.util.List;

import com.microsoft.o365_tasks.R;
import com.microsoft.o365_tasks.auth.AuthCallback;
import com.microsoft.o365_tasks.auth.AuthManager;
import com.microsoft.o365_tasks.data.TaskModel;
import com.microsoft.o365_tasks.sharepoint.SharepointListsClient2;
import com.microsoft.office365.Query;
import com.microsoft.office365.QueryOperations;
import com.microsoft.office365.lists.SPList;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.os.AsyncTask;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.Window;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

public class StartActivity extends Activity {

    private static final String TAG = "StartActivity";
    
    public static final String PARAM_AUTH_IMMEDIATE = "auth_immediate";
  
    private AuthManager mAuthManager;
    private TasksApplication mApplication;

    private TextView mStatusLabel;
    private Button mLoginButton;

    
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        
        requestWindowFeature(Window.FEATURE_NO_TITLE);
        setContentView(R.layout.activity_start);
        
        mApplication = (TasksApplication) getApplication();
        mAuthManager = mApplication.getAuthManager();
        
        //controls
        mStatusLabel = (TextView) findViewById(R.id.status_label);
        mLoginButton = (Button) findViewById(R.id.login_button);
        mLoginButton.setOnClickListener(new View.OnClickListener() {
            @Override public void onClick(View v) {
                start();
            }
        });
        
        reset();

        if (getIntent().getBooleanExtra(PARAM_AUTH_IMMEDIATE, false)) {
            
            start();
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        //Handle authentication completion
        mAuthManager.onActivityResult(requestCode, resultCode, data);
    }

    private void showToast(int resourceId) {
        Toast.makeText(this, resourceId, Toast.LENGTH_SHORT).show();
    }
    
    private void reset() {
        
        //update UI
        mLoginButton.setVisibility(View.VISIBLE);
        mStatusLabel.setVisibility(View.GONE);
    }

    private void start() {
        
        //update UI
        mLoginButton.setVisibility(View.GONE);
        mStatusLabel.setVisibility(View.VISIBLE);
        mStatusLabel.setText(R.string.splash_loading);
        
        //Start authentication procedure
        mAuthManager.forceAuthenticate(this, new AuthCallback() {

            @Override 
            public void onSuccess() {
                startInitProcedure();
            }

            @Override
            public void onFailure(String errorDescription) {
                mStatusLabel.setText(R.string.splash_auth_error);
                launchRetryDialog(errorDescription);
            }

            @Override
            public void onCancelled() {
                reset();
            }

            private void launchRetryDialog(String errorDescription) {
                new AlertDialog.Builder(StartActivity.this)
                    .setTitle(R.string.dialog_auth_failed_title)
                    .setMessage(errorDescription)
                    .setPositiveButton(R.string.label_retry, new DialogInterface.OnClickListener() {
                        @Override public void onClick(DialogInterface dialog, int which) {
                            reset();
                            start();
                        }
                    })
                    .setNegativeButton(R.string.label_cancel, new DialogInterface.OnClickListener() {
                        @Override public void onClick(DialogInterface dialog, int which) {
                            
                        }
                    })
                    .create()
                    .show();
            }
        });
    }

    private void startInitProcedure() {
        
        AsyncTask<Void, Void, Boolean> task = new AsyncTask<Void, Void, Boolean>() {

            @Override
            protected void onPreExecute() {
                showToast(R.string.splash_initializing);
            }
            
            @Override
            protected Boolean doInBackground(Void... ignored) {
                //Executed on background thread
                SharepointListsClient2 listsClient = mApplication.createListsClient();
                
                try {
                    //Does the list already exist?
                    Query query = QueryOperations.field("Title").eq(Constants.SHAREPOINT_LIST_NAME)
                                                 .select("Id");
                    
                    List<SPList> results = listsClient.getLists(query).get();
                    
                    if (results.size() == 1) {
                        //all done - continue
                        return true;
                    }

                    //Create lists in sharepoint
                    //http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splisttemplatetype.aspx
                    final String TASK_LIST = "107";
                    SPList list = listsClient.createList(Constants.SHAREPOINT_LIST_NAME, TASK_LIST).get();
                    
                    //Populate the list with some data
                    TaskModel[] tasks = DemoDataUtil.createDefaultTasks();
                    
                    for (final TaskModel task : tasks) {
                        listsClient.insertListItem(task.getListItem(), list).get();
                    }
                    
                    return true;
                }
                catch (Exception e) {
                    Log.e(TAG, "Error while creating lists", e);
                }
                
                return false;
            }
            
            @Override
            protected void onPostExecute(Boolean success) {

                //Executed on UI thread
                if (!success.booleanValue()) {
                    showToast(R.string.splash_initialization_error);
                    mStatusLabel.setText(R.string.splash_initialization_error);
                }
                else {
                    launchNextActivity();
                }
            }
        };
        
        task.execute();
    }

    private void launchNextActivity() {

        //launch the default next Activity
        final Intent intent = new Intent(this, ListTasksActivity.class);

        startActivity(intent);
        finish();
    }
}
