package com.microsoft.o365_tasks;

import java.util.Date;

import com.microsoft.o365_tasks.controls.DatePickerControl;
import com.microsoft.o365_tasks.data.TaskListItemDataSource;
import com.microsoft.o365_tasks.data.TaskModel;
import com.microsoft.o365_tasks.tasks.ProgressDialogAsyncTask;
import com.microsoft.o365_tasks.utils.AuthUtil;
import com.microsoft.o365_tasks.utils.auth.DefaultAuthHandler;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.os.Bundle;
import android.text.TextUtils;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.Toast;

public class EditTaskActivity extends Activity {
    
    private static final String TAG = "EditTaskActivity";

    public static final String PARAM_TASK_ID = "TaskId";
    
    private static final int CREATING_TASK_ID = -1;
    
    private TasksApplication mApplication;
    private int mTaskId;

    private EditText mTitleText;
    private DatePickerControl mStartDatePicker;
    private DatePickerControl mDueDatePicker;
    private CheckBox mCompletedCheckbox;

    private Button mOkButton;
    private Button mCancelButton;
    
    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_edit_task);

        mApplication = (TasksApplication) getApplication();
        mApplication.getAuthManager();
        
        mTitleText = (EditText) findViewById(R.id.title_text);
        mStartDatePicker = (DatePickerControl) findViewById(R.id.start_date_picker);
        mDueDatePicker = (DatePickerControl) findViewById(R.id.due_date_picker);
        mCompletedCheckbox = (CheckBox) findViewById(R.id.completed_checkbox);
        
        mOkButton = (Button) findViewById(R.id.ok_button);
        mOkButton.setOnClickListener(new OnClickListener() {
            @Override public void onClick(View v) {
                if (validate()) {
                    ensureAuthenticated(new Runnable() {
                        @Override public void run() {
                            saveChangesAndFinish();     
                        }
                    });
                }
            }
        });
        
        mCancelButton = (Button) findViewById(R.id.cancel_button);
        mCancelButton.setOnClickListener(new OnClickListener() {
            @Override public void onClick(View v) {
                //nothing to do - finish the activity
                setResult(RESULT_OK);
                finish();
            }
        });
        
        //read activity arguments
        mTaskId = getIntent().getIntExtra(PARAM_TASK_ID, CREATING_TASK_ID);
        
        start();
    }
   
    //#### Behaviour functions ####

    private void start() {
        if (mTaskId == CREATING_TASK_ID){
            //Creating a new task...
            populateView(createNewTask());
        }
        else {
            //Loading an existing task...
            ensureAuthenticated(new Runnable() {
                @Override public void run() {
                    loadTaskDetails();
                }
            });
        }
    }

    private TaskModel createNewTask() {
        final Date today = new Date();
        final TaskModel model = new TaskModel();
        model.setStartDate(today);
        return model;
    }

    private void populateView(TaskModel model) {
        mTitleText.setText(model.getTitle());
        mStartDatePicker.setDateValue(model.getStartDate());
        mDueDatePicker.setDateValue(model.getDueDate());
        mCompletedCheckbox.setChecked(model.getCompleted());
    }

    private TaskModel createTaskModelFromView() {
        TaskModel model = new TaskModel();
        model.setTitle(mTitleText.getText().toString());
        model.setDueDate(mDueDatePicker.getDateValue());
        model.setStartDate(mStartDatePicker.getDateValue());
        model.setCompleted(mCompletedCheckbox.isChecked());
        return model;
    }

    private boolean validate() {
        String titleText = mTitleText.getText().toString();
        if (TextUtils.isEmpty(titleText)) {
            mTitleText.setError(getString(R.string.edit_task_validation_title_required));
            mTitleText.requestFocus();
            return false;
        }
        return true;
    }
    
    //#### Options menu ####

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.edit_task_options, menu);
        return true;
    }
    
    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();
        switch (id) {
        case R.id.action_delete:
            optionsActionDeleteTask();
            return true;
        }
        return super.onOptionsItemSelected(item);
    }
    
    private void optionsActionDeleteTask() {
        //Launch confirmation dialog
        new AlertDialog.Builder(this)
            .setTitle(R.string.dialog_confirm_title)
            .setMessage(R.string.dialog_delete_confirm_message)
            .setNegativeButton(R.string.label_cancel, null)
            .setPositiveButton(R.string.label_delete, new DialogInterface.OnClickListener() {
                @Override public void onClick(DialogInterface dialog, int which) {
                    
                    //user has confirmed
                    ensureAuthenticated(new Runnable() {
                        @Override public void run() {
                            deleteTask();
                        }
                    });
                    
                }
            })
            .create()
            .show();
    }

    
    //#### Sharepoint api calls ####

    protected void deleteTask() {
       //Launch a background task to delete the current task
       final String message = getString(R.string.edit_task_working);
       new ProgressDialogAsyncTask<Void>(this, message) {
            /**
             * Executes on a background thread. We can safely block here.
             */
            @Override
            protected Void doInBackground(Void... params) {
                try {
                    new TaskListItemDataSource(mApplication).deleteTask(mTaskId);
                }
                catch (Exception ex) {
                    Log.e(TAG, "Error deleting task", ex);
                }
                return null;
            }

            /**
             * Executes on the UI thread after the background thread completes.
             */
            @Override
            protected void onResult(Void result) {
                //all done - this activity is finished
                setResult(RESULT_OK);
                finish();
            }
       }
       .execute();
    }

    private void loadTaskDetails() {
        //Launch a background task to retrieve the current task details 
        final String message = getString(R.string.edit_task_working);
        new ProgressDialogAsyncTask<TaskModel>(this, message) {
            
            /**
             * Executes on a background thread. We can safely block here.
             */
            @Override
            protected TaskModel doInBackground(Void... params) {
                try {
                    return new TaskListItemDataSource(mApplication).getTask(mTaskId);
                }
                catch (Exception ex) {
                    Log.e(TAG, "Error retrieving task", ex);
                    return null;
                }
            }
            
            /**
             * Executes on the UI thread after the background thread completes.
             */
            @Override
            protected void onResult(TaskModel result) {
                if (result != null) {
                    populateView(result);
                }
                else {
                    //something went wrong and we were unable to retrieve a task
                    //alert the user and finish the activity
                    new AlertDialog.Builder(EditTaskActivity.this)
                        .setTitle(R.string.dialog_error_title)
                        .setMessage(R.string.edit_task_error_retrieving_task)
                        .setPositiveButton(R.string.label_continue, new DialogInterface.OnClickListener() {
                            @Override public void onClick(DialogInterface dialog, int which) {
                                //unable to load content - close activity
                                setResult(RESULT_CANCELED);
                                finish();
                            }
                        })
                        .create()
                        .show();
                }
            }
        }
        .execute();
    }

    private void saveChangesAndFinish() {
        //Launch a background task to save these changes
        final String message = getString(R.string.edit_task_working);
        new ProgressDialogAsyncTask<Boolean>(this, message) {

            /**
             * Executes on a background thread. We can safely block here.
             */
            protected Boolean doInBackground(Void... params) {
                try {
                    final TaskModel model = createTaskModelFromView();
                    final TaskListItemDataSource source = new TaskListItemDataSource(mApplication);
                    if (mTaskId == CREATING_TASK_ID) {
                        source.createTask(model);
                        //sleep a moment while the sharepoint item is created
                        Thread.sleep(1500);
                    }
                    else {
                        model.setId(mTaskId);
                        source.updateTask(model);
                    }
                    return true;
                }
                catch (Exception ex) {
                    Log.e(TAG, "Error retrieving task", ex);
                    return false;
                }
            }

            /**
             * Executes on the UI thread after the background thread completes.
             */
            @Override
            protected void onResult(Boolean result) {
                if (result.booleanValue()) {
                    setResult(RESULT_OK);
                    finish();
                }
                else {
                    Toast.makeText(mApplication, R.string.edit_task_error_updating_task, Toast.LENGTH_LONG).show();
                }
            }
        }
        .execute();
    }


    //#### Utility functions ####
    
    private void ensureAuthenticated(final Runnable then) {
        AuthUtil.ensureAuthenticated(this, new DefaultAuthHandler(this) {
            @Override public void onSuccess() {
                then.run();
            }
        });
    }
}
