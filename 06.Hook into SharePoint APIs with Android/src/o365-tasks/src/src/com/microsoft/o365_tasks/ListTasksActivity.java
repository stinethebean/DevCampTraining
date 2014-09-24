package com.microsoft.o365_tasks;

import java.text.DateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import com.microsoft.o365_tasks.R;
import com.microsoft.o365_tasks.data.TaskListItemDataSource;
import com.microsoft.o365_tasks.data.TaskModel;
import com.microsoft.o365_tasks.tasks.ProgressDialogAsyncTask;
import com.microsoft.o365_tasks.utils.AuthUtil;
import com.microsoft.o365_tasks.utils.ViewUtil;
import com.microsoft.o365_tasks.utils.auth.DefaultAuthHandler;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.LayoutInflater;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.ViewGroup;
import android.widget.AdapterView;
import android.widget.BaseAdapter;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;
import android.widget.AdapterView.OnItemClickListener;
import android.widget.ListView;

import android.app.AlertDialog;
import android.content.DialogInterface;
import android.view.ContextMenu;
import android.view.ContextMenu.ContextMenuInfo;
import android.widget.AdapterView.AdapterContextMenuInfo;

import com.microsoft.office365.Query;
import android.content.SharedPreferences;

public class ListTasksActivity extends Activity {
    
    private static final String TAG = "ListTasksActivity";
    
    private static final int REQUEST_EDIT_TASK = 1;

    private TasksApplication mApplication;
    private PreferencesWrapper mPreferences;
    
    private ListView mListView;
    
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_list_tasks);
        
        mApplication = (TasksApplication) getApplication();
        mPreferences = new PreferencesWrapper(mApplication.getSharedPreferences("listtasks_prefs", Context.MODE_PRIVATE));

        mListView = (ListView) findViewById(R.id.list);
        mListView.setOnItemClickListener(new OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
                
                TaskModel task = (TaskModel) mListView.getItemAtPosition(position);
                
                Intent launchIntent = new Intent(ListTasksActivity.this, EditTaskActivity.class);
                launchIntent.putExtra(EditTaskActivity.PARAM_TASK_ID, task.getId());
                startActivityForResult(launchIntent, REQUEST_EDIT_TASK);
            }
        });
        
        registerForContextMenu(mListView);
        
        optionsActionRefresh();
    }
    
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        
        if (requestCode == REQUEST_EDIT_TASK && resultCode == Activity.RESULT_OK) {

            optionsActionRefresh();
        }
    }
    
    //#### Context menu ####
    
    @Override
    public void onCreateContextMenu(ContextMenu menu, View v, ContextMenuInfo menuInfo) {

        getMenuInflater().inflate(R.menu.list_tasks_context, menu);
    }

    
    @Override
    public boolean onContextItemSelected(MenuItem item) {
        AdapterContextMenuInfo info = (AdapterContextMenuInfo) item.getMenuInfo();
        switch (item.getItemId()) {
        case R.id.action_delete:
            contextActionDelete(info);
            return true;
        }
        
        return super.onContextItemSelected(item);
    }


    private void contextActionDelete(AdapterContextMenuInfo info) {
        final TaskModel task = (TaskModel) mListView.getItemAtPosition(info.position);

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
                            
                            deleteTask(task);
                        }
                    });
                }
            })
            .create()
            .show();
    }

    //#### Options menu ####

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.list_tasks_options, menu);
        
        menu.findItem(R.id.action_filter_completed).setChecked(mPreferences.getFilterCompleted());
        
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();
        switch (id) {
        case R.id.action_clear_token:
            optionsActionClearToken();
            return true;
            
        case R.id.action_add_new:
            optionsActionAddNewTask();
            return true;
            
        case R.id.action_refresh:
            optionsActionRefresh();
            return true;
            
        case R.id.action_filter_completed:
            optionsActionFilterCompleted(item);
            return true;
        }
        return super.onOptionsItemSelected(item);
    }

    private void optionsActionClearToken() {
        mApplication.getAuthManager().forceExpireToken();
    }

    private void optionsActionAddNewTask() {
        Intent launchIntent = new Intent(ListTasksActivity.this, EditTaskActivity.class);
        startActivityForResult(launchIntent, REQUEST_EDIT_TASK);
    }

    private void optionsActionRefresh() {
        ensureAuthenticated(new Runnable() {
            @Override public void run() {
                refresh();
            }
        });
    }

    private void optionsActionFilterCompleted(MenuItem item) {
        
        boolean flag = !item.isChecked();
        item.setChecked(flag);
        mPreferences.setFilterCompleted(flag);
        
        //refresh
        optionsActionRefresh();
    }
    
    //#### Sharepoint api calls ####

    private void refresh() {
        
        //start task to retreive list items
        final String message = getString(R.string.list_tasks_retrieving_tasks);
        new ProgressDialogAsyncTask<List<TaskModel>>(this, message) {

            /**
             * Executes on a background thread. We can safely block here.
             */
            @Override
            protected List<TaskModel> doInBackground(Void... params) {
                try {
                    
                    Query query = new Query();
                    
                    if (mPreferences.getFilterCompleted()) {
                        query.field("PercentComplete").lt(TaskModel.COMPLETED_MAX);
                    }
                    
                    return new TaskListItemDataSource(mApplication).getTasksByQuery(query);
                }
                catch (Exception ex) {
                    Log.e(TAG, "Error retrieving task list", ex);
                    return null;
                }
            }

            /**
             * Executes on the UI thread after the background thread completes.
             */
            @Override
            protected void onResult(List<TaskModel> results) {
                if (results == null) {
                    //error retrieving results - alert the user
                    Toast.makeText(mApplication, R.string.list_tasks_error_retrieving_tasks, Toast.LENGTH_LONG).show();
                    results = Collections.emptyList();
                }
                //render the list items onto the view
                final TaskListAdapter adapter = new TaskListAdapter(ListTasksActivity.this, results);
                mListView.setAdapter(adapter);
                adapter.notifyDataSetChanged();
            }
        }
        .execute();
    }

    protected void deleteTask(final TaskModel model) {
       //Launch a background task to delete the current task
       final String message = getString(R.string.edit_task_working);
       new ProgressDialogAsyncTask<Void>(this, message) {
            /**
             * Executes on a background thread. We can safely block here.
             */
            @Override
            protected Void doInBackground(Void... params) {
                try {
                    new TaskListItemDataSource(mApplication).deleteTask(model);
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
                //all done - refresh the list
                refresh();
            }
       }
       .execute();
    }

    //#### Utility functions ####
    
    private void ensureAuthenticated(final Runnable then) {
        AuthUtil.ensureAuthenticated(this, new DefaultAuthHandler(this) {
            @Override public void onSuccess() {
                //The access token has been refreshed, or is still valid
                //It is safe to continue!
                then.run();
            }
        });
    }
    
    //#### Utility classes ####
    
    /**
     * The TaskListAdapter adapts a list of TaskModels into views.
     * This class specifically is used with a ListView which will repeatedly invoke getView()
     * for each item in order to build the list.
     */
    public class TaskListAdapter extends BaseAdapter {

        private Context mContext;
        private LayoutInflater mInflater;
        private DateFormat mDateFormat;
        
        private List<TaskModel> mData;

        public TaskListAdapter (Context context, List<TaskModel> data){
            
            assert context != null;
            assert data != null;
            
            mContext = context;
            mInflater = (LayoutInflater) mContext.getSystemService(Context.LAYOUT_INFLATER_SERVICE);
            mDateFormat = DateFormat.getDateInstance(DateFormat.LONG);
            
            mData = data;
        }

        /**
         * Creates or refreshes a View for the TaskModel item at the given list position. 
         */
        @Override 
        public View getView(int position, View convertView, ViewGroup parent) {
            View view = ViewUtil.prepareView(mInflater, R.layout.task_list_item, convertView, parent);

            final TextView title          = (TextView) ViewUtil.findChildView(view, R.id.title);
            final ImageView completedFlag = (ImageView) ViewUtil.findChildView(view, R.id.completed_flag);
            final View startDateContainer = (View)     ViewUtil.findChildView(view, R.id.startDate);
            final TextView startDateLabel = (TextView) ViewUtil.findChildView(view, R.id.startDate_label);
            final View dueDateContainer   = (View)     ViewUtil.findChildView(view, R.id.dueDate);
            final TextView dueDateLabel   = (TextView) ViewUtil.findChildView(view, R.id.dueDate_label);
            
            final TaskModel model = (TaskModel) getItem(position);
            
            title.setText(model.getTitle());
            
            Date startDateValue = model.getStartDate();
            updateDateView(startDateLabel, startDateContainer, startDateValue);
            
            Date dueDateValue = model.getDueDate();
            updateDateView(dueDateLabel, dueDateContainer, dueDateValue);
            
            completedFlag.setImageResource(model.getCompleted() ? R.drawable.ic_checkbox_checked : R.drawable.ic_checkbox_unchecked);
            
            return view;
        }
        
        private void updateDateView(final TextView label, final View container, final Date dateValue) {
            if (dateValue == null){
                container.setVisibility(View.GONE);
            }
            else {
                container.setVisibility(View.VISIBLE);
                label.setText(mDateFormat.format(dateValue));
            }
        }

        @Override
        public int getCount() {
            return mData.size();
        }

        @Override
        public Object getItem(int position) {
            return mData.get(position);
        }

        @Override
        public long getItemId(int position) {
            return position;
        }
        
    }
    
    private static class PreferencesWrapper {
        
        private static final String PREFS_FILTER_COMPLETED = "filter_completed";
        
        private SharedPreferences mPreferences;
                
        public PreferencesWrapper(SharedPreferences preferences) {
            mPreferences = preferences;
        }

        public boolean getFilterCompleted() {
            return mPreferences.getBoolean(PREFS_FILTER_COMPLETED, false);
        }
        
        public void setFilterCompleted(boolean completed) {
            mPreferences.edit()
                        .putBoolean(PREFS_FILTER_COMPLETED, completed)
                        .apply();
        }
    }
}
