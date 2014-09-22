package com.microsoft.o365_tasks.utils;

public class DateParseException extends Exception {

    private static final long serialVersionUID = 1L;
    
    private final String mInput;
    private final int mPosition;

    public DateParseException(String message, String input, int position) {
        super(String.format("%s - position: %d, input: %s", message, position, input));
        mInput = input;
        mPosition = position;
    }

    public String getInput() {
        return mInput;
    }

    public int getPosition() {
        return mPosition;
    }
}
