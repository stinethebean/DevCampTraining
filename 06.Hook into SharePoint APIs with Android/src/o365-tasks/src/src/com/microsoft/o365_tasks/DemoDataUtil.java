package com.microsoft.o365_tasks;

import java.util.Calendar;
import java.util.Date;

import com.microsoft.o365_tasks.data.TaskModel;

public class DemoDataUtil {

    private static TaskModel createTask(String title, Date startDate, Date dueDate, boolean completed) {
        TaskModel m = new TaskModel();
        m.setTitle(title);
        m.setStartDate(startDate);
        m.setDueDate(dueDate);
        m.setCompleted(completed);
        return m;
    }

    public static TaskModel[] createDefaultTasks() {
        final Calendar c = Calendar.getInstance();
        final Date now = c.getTime();
        c.add(Calendar.DAY_OF_MONTH, 5);
        final Date then = c.getTime();
        final TaskModel[] tasks = {
            createTask("Pick up bread, milk", null, null, false),
            createTask("Client meeting prep", now, then, true),
            createTask("Wash the car", null, then, true),
            createTask("Book vacation", now, then, false),
            createTask("Pick up drycleaning", now, null, false),
            createTask("Unclog the sink", null, null, true),
            createTask("Call mom", now, then, false)
        };
        return tasks;
    }
}