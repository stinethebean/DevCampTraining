package com.microsoft.o365_tasks.controls;

import java.util.Calendar;
import java.util.Date;

import android.app.DatePickerDialog;
import android.app.DatePickerDialog.OnDateSetListener;
import android.content.Context;
import android.content.DialogInterface;
import android.content.DialogInterface.OnCancelListener;
import android.widget.DatePicker;

public class DatePickerControlUtil {
    
    /**
     * Because of a tooling bug, using OnDateSetListener within the control class causes the Android
     * layout editor to report errors (java.lang.ClassNotFoundException: android.app.DatePickerDialog$OnDateSetListener).
     * 
     * This function was extracted from the DatePickerControl to work around that bug.
     * 
     * @param context
     * @param control
     */
    public static void launchDatePicker(Context context, final DatePickerControl control, boolean cancelable) {
        final Date initialDate = control.getDateValue();
        final Calendar c = Calendar.getInstance();
        if (initialDate != null) {
            c.setTime(initialDate);
        }
        OnDateSetListener listener = new OnDateSetListener() {
            @Override public void onDateSet(DatePicker view, int year, int monthOfYear, int dayOfMonth) {
                c.set(Calendar.YEAR, year);
                c.set(Calendar.MONTH, monthOfYear);
                c.set(Calendar.DAY_OF_MONTH, dayOfMonth);
                control.setDateValue(c.getTime());
            }
        };
        DatePickerDialog dialog = new DatePickerDialog(
            context, 
            listener, 
            c.get(Calendar.YEAR), 
            c.get(Calendar.MONTH), 
            c.get(Calendar.DAY_OF_MONTH)
        );
        dialog.setCancelable(cancelable);
        if (cancelable) {
            dialog.setOnCancelListener(new OnCancelListener() {
                @Override public void onCancel(DialogInterface dialog) {
                    //If the user dismissed the dialog without selecting "OK" then undo any changes they made
                    control.setDateValue(initialDate);
                }
            });
        }
        dialog.show();
    }
    
}
