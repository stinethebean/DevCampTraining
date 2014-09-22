package com.microsoft.o365_tasks.data;

import java.util.Date;

import com.microsoft.o365_tasks.utils.SPUtil;
import com.microsoft.office365.lists.SPListItem;

public class TaskModel {
    
    public static final double COMPLETED_MIN = 0;
    public static final double COMPLETED_MAX = 1;
    public static final String[] SELECT_FIELDS = { 
        "Id", "Title", "PercentComplete", "StartDate", "DueDate"
    };
    
    private SPListItem mListItem;

    public TaskModel() {
        mListItem = new SPListItem();
    }
    
    public TaskModel(SPListItem listItem) {
        assert listItem != null;
        mListItem = listItem;
    }

    public SPListItem getListItem() {
        return mListItem;
    }
    
    /* ------ Id ------ */
    public int getId() {
        return SPUtil.safeInt(mListItem, "Id", 0);
    }

    public void setId(int id) {
        mListItem.setData("Id", SPUtil.jsonValue(id));
    }
    
    /* ------ Task Name ------ */
    public String getTitle() {
        return SPUtil.safeString(mListItem, "Title", "");
    }

    public void setTitle(String taskName) {
        mListItem.setData("Title", SPUtil.jsonValue(taskName));
    }
    
    /* ------ % Complete ------ */
    public double getPercentComplete() {
        return SPUtil.safeDouble(mListItem, "PercentComplete", 0);
    }

    public void setPercentComplete(double percentComplete) {
        mListItem.setData("PercentComplete", SPUtil.jsonValue(percentComplete));
    }
    
    public boolean getCompleted() {
        return getPercentComplete() >= COMPLETED_MAX;
    }
    
    public void setCompleted(boolean value){
        setPercentComplete(value ? COMPLETED_MAX : COMPLETED_MIN);
    }
    
    /* ------ Start Date ------ */
    public Date getStartDate() {
        return SPUtil.safeDate(mListItem, "StartDate", null);
    }

    public void setStartDate(Date startDate) {
        mListItem.setData("StartDate", SPUtil.jsonValue(startDate));
    }
    
    /* ------ Due Date ------ */
    public Date getDueDate() {
        return SPUtil.safeDate(mListItem, "DueDate", null);
    }

    public void setDueDate(Date dueDate) {
        mListItem.setData("DueDate", SPUtil.jsonValue(dueDate));
    }
}
