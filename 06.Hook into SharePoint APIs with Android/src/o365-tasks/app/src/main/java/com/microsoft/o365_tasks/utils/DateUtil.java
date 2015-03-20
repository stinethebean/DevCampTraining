package com.microsoft.o365_tasks.utils;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.TimeZone;

public class DateUtil {

    private static final String UTC_TIMEZONE_ID = "UTC";

    private static String zeroPad(int number, int padLength) {

        String str = Integer.toString(number);

        if (str.length() > padLength)
            return str;

        char[] zeroes = new char[padLength - str.length()];
        for (int i = 0; i < zeroes.length; i++) {
            zeroes[i] = '0';
        }

        return new String(zeroes) + str;
    }

    /**
     * ISO8601 date/time parser based on the parser from Json.NET.
     * 
     * https://github.com/JamesNK/Newtonsoft.Json/blob/master/Src/Newtonsoft.Json/Utilities/DateTimeParser.cs
     * 
     * This parser can handle optional date/time parts and will happily parse up to 7 digits of millisecond 
     * precision, as can be the case for ISO8601 strings coming from .NET applications.
     * 
     * Unfortunaly the Java Date class only supports up to 999ms of precision, so any additional precision is lost.
     * 
     */
    private static class Iso8601Parser
    {
        private static final int CONSUME_FAIL = -1;
        private static final int MAX_FRACTION_DIGITS = 7;
        private static final int TICKS_PER_MILLISECOND = 10000;
        private static final int[] POWERS_OF_TEN = new int[] { -1 /* not used */, 10, 100, 1000, 10000, 100000, 1000000 };

        private char[] mInput;
        private int mPosition;

        private int mYear;
        private int mMonth;
        private int mDayOfMonth;
        private int mHour;
        private int mMinute;
        private int mSecond;
        private int mFraction;
        private String mTimeZoneId;

        public Date parseIso8601(String dateString) throws DateParseException {

            mInput = dateString.toCharArray();
            mPosition = 0;

            mYear = 0;
            mMonth = 1;
            mDayOfMonth = 1;
            mHour = 0;
            mMinute = 0;
            mSecond = 0;
            mFraction = 0;
            mTimeZoneId = UTC_TIMEZONE_ID;

            parseDateComponent();

            if (tryConsumeChar('T') || tryConsumeChar('t')) {

                parseTimeComponent();

                if (tryConsumeChar('.')) {

                    parseMillisecondComponent();
                }

                parseTimeZoneComponent();

            }

            //assert this is the end of the string
            if (mPosition < mInput.length)
                throw makeParserError("Unexpected character '" + mInput[mPosition] + "'");

            TimeZone timeZone = TimeZone.getTimeZone(mTimeZoneId);

            if (!timeZone.getID().equalsIgnoreCase(mTimeZoneId)) {
                throw makeParserError("Invalid time zone: " + mTimeZoneId);
            }

            Calendar cal = new GregorianCalendar(timeZone);
            cal.setLenient(false);

            cal.set(Calendar.YEAR, mYear);
            cal.set(Calendar.MONTH, mMonth - 1);
            cal.set(Calendar.DAY_OF_MONTH, mDayOfMonth);
            cal.set(Calendar.HOUR_OF_DAY, mHour);
            cal.set(Calendar.MINUTE, mMinute);
            cal.set(Calendar.SECOND, mSecond);
            cal.set(Calendar.MILLISECOND, mFraction / TICKS_PER_MILLISECOND);

            return cal.getTime();
        }

        private void parseDateComponent() throws DateParseException {

            boolean ok;

            //Parse year component
            ok = (mYear = consume4Digits()) != CONSUME_FAIL
                    && 1 <= mYear;

            if (!ok) throw makeParserError("Invalid year component");

            consumeChar('-');

            //Parse month component
            ok = (mMonth = consume2Digits()) != CONSUME_FAIL
                    && 1 <= mMonth
                    && mMonth <= 12;

            if (!ok) throw makeParserError("Invalid month component");

            consumeChar('-');

            //Parse day component
            ok = (mDayOfMonth = consume2Digits()) != CONSUME_FAIL
                    && 1 <= mDayOfMonth; //check start of range

            if (!ok) throw makeParserError("Invalid day component");
        }

        private void parseTimeComponent() throws DateParseException {

            boolean ok;

            //Parse hour component
            ok = (mHour = consume2Digits()) != CONSUME_FAIL
                    && mHour < 24;

            if (!ok) throw makeParserError("Invalid hour component");

            consumeChar(':');

            //Parse minutes component
            ok = (mMinute = consume2Digits()) != CONSUME_FAIL
                    && mMinute < 60;

            if (!ok) throw makeParserError("Invalid minute component");

            consumeChar(':');

            //Parse seconds component
            ok = (mSecond = consume2Digits()) != CONSUME_FAIL
                    && mSecond < 60;

            if (!ok) throw makeParserError("Invalid second component");
        }

        private void parseMillisecondComponent() throws DateParseException {
            int numberOfDigits = 0;

            while (mPosition < mInput.length && numberOfDigits < MAX_FRACTION_DIGITS)
            {
                int digit = mInput[mPosition] - '0';
                if (digit < 0 || digit > 9)
                    break;

                mFraction = (mFraction * 10) + digit;

                numberOfDigits++;
                mPosition++;
            }

            if (numberOfDigits < MAX_FRACTION_DIGITS) {

                if (numberOfDigits == 0)
                    throw makeParserError("Invalid millisecond component");

                mFraction *= POWERS_OF_TEN[MAX_FRACTION_DIGITS - numberOfDigits];
            }
        }

        private void parseTimeZoneComponent() throws DateParseException {

            if (tryConsumeChar('Z') || tryConsumeChar('z')) {

                mTimeZoneId = UTC_TIMEZONE_ID;

            }
            else {
                boolean isPositive = tryConsumeChar('+');

                if (!isPositive && !tryConsumeChar('-')) {

                    //No time zone component! bail out
                    return;
                }

                boolean ok;

                //Hour component
                int zoneHour = 0;

                ok = (zoneHour = consume2Digits()) != CONSUME_FAIL
                        && zoneHour < 24;

                if (!ok) throw makeParserError("Invalid zone hour component");

                tryConsumeChar(':'); //colon is optional

                //Minute component
                int zoneMinute = 0;

                ok = (zoneMinute = consume2Digits()) != CONSUME_FAIL
                        && zoneMinute < 60;

                if (!ok) throw makeParserError("Invalid zone minute component");

                mTimeZoneId = "GMT" + (isPositive ? '+' : '-') + zeroPad(zoneHour, 2) + ':' + zeroPad(zoneMinute, 2);
            }
        }

        private DateParseException makeParserError(String message) {
            return new DateParseException(message, new String(mInput), mPosition);
        }

        /**
         * Consumes four digits from the input.
         * @return The decimal value or CONSUME_FAIL
         */
        private int consume4Digits()
        {
            if (mPosition + 3 < mInput.length)
            {
                int digit1 = mInput[mPosition + 0] - '0';
                int digit2 = mInput[mPosition + 1] - '0';
                int digit3 = mInput[mPosition + 2] - '0';
                int digit4 = mInput[mPosition + 3] - '0';

                if (0 <= digit1 && digit1 < 10 &&
                    0 <= digit2 && digit2 < 10 &&
                    0 <= digit3 && digit3 < 10 &&
                    0 <= digit4 && digit4 < 10) {

                    mPosition += 4;
                    return (((((digit1 * 10) + digit2) * 10) + digit3) * 10) + digit4;
                }
            }

            return CONSUME_FAIL;
        }

        /**
         * Consumes two decimal digits from the input.
         * @return The decimal value or CONSUME_FAIL
         */
        private int consume2Digits()
        {
            if (mPosition + 1 < mInput.length)
            {
                int digit1 = mInput[mPosition + 0] - '0';
                int digit2 = mInput[mPosition + 1] - '0';

                if (0 <= digit1 && digit1 < 10 &&
                    0 <= digit2 && digit2 < 10) {

                    mPosition += 2;
                    return (digit1 * 10) + digit2;
                }
            }

            return CONSUME_FAIL;
        }

        /**
         * Attempts to consume a character from the input.
         * Throws if unsuccessful.
         * @param ch The character to consume
         * @throws DateParseException
         */
        private void consumeChar(char ch) throws DateParseException {
            if (!tryConsumeChar(ch)) {
                throw makeParserError("Expected '" + ch + "'");
            }
        }

        /**
         * Attempts to consume a character from the input.
         * The input position is only advanced when successful.
         * @param ch The character to consume
         * @return true if successful
         */
        private boolean tryConsumeChar(char ch)
        {
            if (mPosition < mInput.length && mInput[mPosition] == ch) {
                mPosition += 1;
                return true;
            }

            return false;
        }
    }

    public static Date parseIso8601(String s) throws DateParseException {

        Iso8601Parser parser = new Iso8601Parser();

        return parser.parseIso8601(s);
    }

    public static String formatIso8601(Date date) {
        return formatIso8601(date, TimeZone.getTimeZone(UTC_TIMEZONE_ID));
    }

    public static String formatIso8601(Date date, TimeZone tz) {

        Calendar calendar = new GregorianCalendar(tz, Locale.US);
        calendar.setTime(date);

        //Any formatted string should fit within 35 characters...
        StringBuilder sb = new StringBuilder(35);

        sb.append(zeroPad(calendar.get(Calendar.YEAR), 4));
        sb.append('-');
        sb.append(zeroPad(calendar.get(Calendar.MONTH) + 1, 2));
        sb.append('-');
        sb.append(zeroPad(calendar.get(Calendar.DAY_OF_MONTH), 2));
        sb.append('T');
        sb.append(zeroPad(calendar.get(Calendar.HOUR_OF_DAY), 2));
        sb.append(':');
        sb.append(zeroPad(calendar.get(Calendar.MINUTE), 2));
        sb.append(':');
        sb.append(zeroPad(calendar.get(Calendar.SECOND), 2));
        sb.append('.');
        sb.append(zeroPad(calendar.get(Calendar.MILLISECOND), 3)); //maximum accuracy of Calendar is 999 ms

        int offset = tz.getOffset(calendar.getTimeInMillis());
        if (offset == 0) {
            sb.append('Z');
        }
        else {
            int hours = Math.abs((offset / (60 * 1000)) / 60);
            int minutes = Math.abs((offset / (60 * 1000)) % 60);
            sb.append(offset < 0 ? '-' : '+');
            sb.append(zeroPad(hours, 2));
            sb.append(':');
            sb.append(zeroPad(minutes, 2));
        }

        return sb.toString();
    }

}
