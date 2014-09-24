package com.microsoft.o365_tasks.auth;

public interface AuthCallback {
    public void onFailure(String errorDescription);
    public void onCancelled();
    public void onSuccess();
}
