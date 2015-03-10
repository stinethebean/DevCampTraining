package com.microsoft.o365_tasks.controls;

import java.text.DateFormat;
import java.util.Date;

import android.content.Context;
import android.content.res.TypedArray;
import android.util.AttributeSet;
import android.view.LayoutInflater;
import android.view.View;
import android.widget.Button;
import android.widget.ImageButton;
import android.widget.LinearLayout;

import com.microsoft.o365_tasks.R;

public class DatePickerControl extends LinearLayout {

    private Context mContext;
    private DateFormat mDateformat;
    
    private Button mPickerButton;
    private ImageButton mClearButton;
    
    private Date mCurrentDate;
    private boolean mClearable = true;

    public DatePickerControl(Context context) {
        super(context);
        init(context);
    }

    public DatePickerControl(Context context, AttributeSet attrs) {
        super(context, attrs);
        initFromAttributes(context, attrs);
        init(context);
    }

    public DatePickerControl(Context context, AttributeSet attrs, int defStyle) {
        super(context, attrs, defStyle);
        initFromAttributes(context, attrs);
        init(context);
    }

    private void init(Context context) {
        mContext = context;
        mDateformat = android.text.format.DateFormat.getLongDateFormat(mContext);
        
        LayoutInflater inflater = (LayoutInflater) mContext.getSystemService(Context.LAYOUT_INFLATER_SERVICE);
        inflater.inflate(R.layout.component_date_picker_control, this);
        
        mPickerButton = (Button) findViewById(R.id.launch_picker);
        mPickerButton.setOnClickListener(new OnClickListener() {
            @Override public void onClick(View v) {
                DatePickerControlUtil.launchDatePicker(mContext, DatePickerControl.this, mClearable);
            }
        });
        
        mClearButton = (ImageButton) findViewById(R.id.clear_picker);
        mClearButton.setOnClickListener(new OnClickListener() {
            @Override public void onClick(View v) {
                setDateValue(null);
            }
        });
        
        updateView();
    }

    private void initFromAttributes(Context context, AttributeSet attrs) {
        TypedArray a = context.obtainStyledAttributes(attrs, R.styleable.DatePickerControl);
        mClearable = a.getBoolean(R.styleable.DatePickerControl_clearable, true);
        a.recycle();
    }

    private void updateView() {
        String label =
            (mCurrentDate != null) 
                ? mDateformat.format(mCurrentDate)
                : mContext.getString(R.string.component_date_picker_no_date);
                
        mPickerButton.setText(label);
        mClearButton.setVisibility(mClearable ? View.VISIBLE : View.GONE);
    }
    
    public boolean isDateValueSet() {
        return mCurrentDate != null;
    }
    
    public Date getDateValue() {
        return mCurrentDate;
    }
    
    public void setDateValue(Date value) {
        mCurrentDate = value;
        updateView();
    }
    
    public boolean getClearable() {
        return mClearable;
    }
    
    public void setClearable(boolean value) {
        mClearable = value;
        updateView();
    }
}
