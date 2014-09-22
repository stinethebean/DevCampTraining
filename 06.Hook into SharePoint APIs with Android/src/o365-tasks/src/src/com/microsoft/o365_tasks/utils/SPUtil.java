package com.microsoft.o365_tasks.utils;

import java.util.Date;

import org.json.JSONObject;

import android.util.Log;

import com.microsoft.office365.lists.SPListItem;

public class SPUtil {
    
    private static final String TAG = "SPUtil";
    
    public static String safeString(SPListItem listItem, String field, String defaultValue) {
        if (listItem != null) {
            try {
                Object value = listItem.getData(field);
                if (value == null || value.equals(JSONObject.NULL)) {
                    return defaultValue;
                }
                return value.toString();
            }
            catch (Exception ex) {
                Log.w(TAG, "safeString: " + Log.getStackTraceString(ex));
            }
        }
        return defaultValue;
    }
    
    public static Object jsonValue(String value) {
        return value;
    }

    public static int safeInt(SPListItem listItem, String field, int defaultValue) {
        if (listItem != null) {
            try {
                Object value = listItem.getData(field);
                if (value instanceof Double) {
                    return (int) ((Double)value).doubleValue();
                }
                if (value instanceof Integer) {
                    return ((Integer)value).intValue();
                }
            }
            catch (Exception ex) {
                Log.w(TAG, "safeInt: " + Log.getStackTraceString(ex));
            }
        }
        return defaultValue;
    }
    
    public static Object jsonValue(int value) {
        return (Integer) value;
    }

    public static double safeDouble(SPListItem listItem, String field, double defaultValue) {
        if (listItem != null) {
            try {
                Object value = listItem.getData(field);
                if (value instanceof Double) {
                    return ((Double)value).doubleValue();
                }
                if (value instanceof Integer) {
                    return (double) ((Integer)value).intValue();
                }
            }
            catch (Exception ex) {
                Log.w(TAG, "safeDouble: " + Log.getStackTraceString(ex));
            }
        }
        return defaultValue;
    }
    
    public static Object jsonValue(double value) {
        return (Double) value;
    }
    
    public static Date safeDate(SPListItem listItem, String field, Date defaultValue) {
        if (listItem != null) {
            try {
                Object value = listItem.getData(field);
                return DateUtil.parseIso8601(value.toString());
            }
            catch (Exception ex) {
                Log.w(TAG, "safeDate: " + Log.getStackTraceString(ex));
            }
        }
        return defaultValue;
    }

    public static Object jsonValue(Date date) {
        return date == null ? JSONObject.NULL : DateUtil.formatIso8601(date);
    }
}
