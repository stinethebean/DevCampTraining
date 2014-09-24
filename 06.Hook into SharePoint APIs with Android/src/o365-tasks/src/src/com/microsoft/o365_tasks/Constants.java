package com.microsoft.o365_tasks;

public class Constants {

    public static final String AAD_DOMAIN = "YOURDOMAIN.onmicrosoft.com";
    public static final String AAD_CLIENT_ID = "00000000-0000-0000-0000-000000000000";
    public static final String AAD_LOGIN_HINT = "user@" + AAD_DOMAIN;
    public static final String AAD_AUTHORITY = "https://login.windows.net/" + AAD_DOMAIN;
    public static final String AAD_REDIRECT_URL = "http://android/complete";
    
    public static final String SHAREPOINT_URL = "https://YOURDOMAIN.sharepoint.com";
    public static final String SHAREPOINT_SITE_PATH = "";
    public static final String SHAREPOINT_LIST_NAME = "Tasks for Android";

}
