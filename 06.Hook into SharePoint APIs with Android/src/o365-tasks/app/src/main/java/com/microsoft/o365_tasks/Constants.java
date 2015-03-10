package com.microsoft.o365_tasks;

public class Constants {

    public static final String AAD_DOMAIN = "o365performancerecorder.onmicrosoft.com";
    public static final String AAD_CLIENT_ID = "566c7295-9bdf-4d83-81e6-92db6d7ae4b3";
    public static final String AAD_LOGIN_HINT = "admin@" + AAD_DOMAIN;
    public static final String AAD_AUTHORITY = "https://login.windows.net/" + AAD_DOMAIN;
    public static final String AAD_REDIRECT_URL = "http://android/complete";
    
    public static final String SHAREPOINT_URL = "https://o365performancerecorder.sharepoint.com";
    public static final String SHAREPOINT_SITE_PATH = "";
    public static final String SHAREPOINT_LIST_NAME = "Tasks for Android";

}
