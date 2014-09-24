package com.microsoft.o365_tasks.data;

import java.util.ArrayList;
import java.util.List;

import com.microsoft.o365_tasks.Constants;
import com.microsoft.o365_tasks.TasksApplication;
import com.microsoft.o365_tasks.sharepoint.SharepointListsClient2;
import com.microsoft.office365.Query;
import com.microsoft.office365.QueryOperations;
import com.microsoft.office365.lists.SPList;
import com.microsoft.office365.lists.SPListItem;

public class TaskListItemDataSource {

    private TasksApplication mApplication;
    
    private SharepointListsClient2 mClient;

    public TaskListItemDataSource(TasksApplication application) {
        mApplication = application;
    }
    
    private SharepointListsClient2 getListsClient() {
        if (mClient == null) {
            mClient = mApplication.createListsClient();
        }
        return mClient;
    }

    /**
     * Makes a blocking call to the Sharepoint API to execute the given query against the
     * configured Tasks list.
     * 
     * @param query The query to execute. Can be null.
     * @return The resolved list of TaskListItemModels
     * @throws Exception
     */
    public List<TaskModel> getTasksByQuery(Query query) throws Exception {
        
        final ArrayList<TaskModel> items = new ArrayList<TaskModel>();

        if (query == null) {
            query = new Query();
        }
        
        //Limit the query to just the fields we are interested in
        query.select(TaskModel.SELECT_FIELDS);
        
        List<SPListItem> listItems = getListsClient().getListItems(Constants.SHAREPOINT_LIST_NAME, query).get();
        
        for (final SPListItem item : listItems) {
            items.add(new TaskModel(item));
        }
        
        return items;
    }

    /**
     * Makes a blocking call to the Sharepoint API to retrieve the Task list item with
     * the given Id.
     * 
     * @param id The id of the list item to retrieve.
     * @return The resolved item, or null if it was not found.
     * @throws Exception
     */
    public TaskModel getTask(int id) throws Exception {
        
        //Generate the OData query $filter=Id eq (#id)&$top=1
        Query query = QueryOperations.field("Id").eq(id)
                                     .top(1);

        //Limit the query to just the fields we are interested in
        query.select(TaskModel.SELECT_FIELDS);
        
        List<SPListItem> listItems = getListsClient().getListItems(Constants.SHAREPOINT_LIST_NAME, query).get();
        
        if (listItems.size() == 0) {
            return null;
        }
        
        return new TaskModel(listItems.get(0));
    }

    /**
     * Makes a blocking call to the Sharepoint API to update the Task based on the
     * item passed in argument.
     * 
     * @param model
     * @throws Exception
     */
    public void updateTask(TaskModel model) throws Exception {
        
        assert model != null;
        
        SharepointListsClient2 client = getListsClient();
        
        SPList list = client.getList(Constants.SHAREPOINT_LIST_NAME).get();
        
        client.updateListItem(model.getListItem(), list).get();
    }

    /**
     * Makes a blocking call to the Sharepoint API to create a new Task based on the
     * model passed in argument.
     * 
     * @param model
     * @throws Exception
     */
    public void createTask(TaskModel model) throws Exception {
        
        assert model != null;
        
        SharepointListsClient2 client = getListsClient();
        
        SPList list = client.getList(Constants.SHAREPOINT_LIST_NAME).get();
        
        client.insertListItem(model.getListItem(), list);
    }

    /**
     * Makes a blocking call to the Sharepoint API to delete the Task with the given Id.
     * 
     * @param model
     * @throws Exception
     */
    public void deleteTask(int id) throws Exception {
        
        TaskModel model = getTask(id);
        
        if (model != null) {

            deleteTask(model);
        }
    }

    /**
     * Makes a blocking call to the Sharepoint API to delete the Task with the given Id.
     * 
     * @param model
     * @throws Exception
     */
    public void deleteTask(TaskModel model) throws Exception {
        
        assert model != null;

        getListsClient().deleteListItem(model.getListItem(), Constants.SHAREPOINT_LIST_NAME).get();
    }
}
