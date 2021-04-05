package com.github.helloichen.util;

import java.text.DateFormat;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 获取SimpleDateFormat保证线程安全
 * 格式化时间
 *
 * @author Chenwp
 */
public class DateUtils {
    /**
     * 存放不同的日期模板格式的SimpleDateFormat的Map
     */
    private static Map<String, ThreadLocal<SimpleDateFormat>> SDF_FORMATTER_HOLDER = new HashMap<>(16);

    /**
     * 返回一个ThreadLocal的sdf,每个线程只会new一次sdf
     *
     * @param pattern 格式
     * @return SimpleDateFormat
     */
    private static SimpleDateFormat getSimpleDateFormat(final String pattern) {
        ThreadLocal<SimpleDateFormat> threadLocal = SDF_FORMATTER_HOLDER.get(pattern);

        // 此处的双重判断和同步是为了防止sdfMap这个单例被多次put重复的sdf
        if (threadLocal == null) {
            synchronized (DateUtils.class) {
                threadLocal = SDF_FORMATTER_HOLDER.get(pattern);
                // 只有Map中还没有这个pattern的sdf才会生成新的sdf并放入map
                if (threadLocal == null) {
                    // 这里是关键,使用ThreadLocal<SimpleDateFormat>替代原来直接new SimpleDateFormat
                    threadLocal = ThreadLocal.withInitial(() -> new SimpleDateFormat(pattern));
                    SDF_FORMATTER_HOLDER.put(pattern, threadLocal);
                }
            }
        }
        return threadLocal.get();
    }

    /**
     * 将字符串(格式符合规范)转换成Date
     *
     * @param value   需要转换的字符串
     * @param pattern 日期格式
     * @return Date
     */
    public static Date string2Date(String value, String pattern) {
        if (value == null || "".equals(value)) {
            return null;
        }
        SimpleDateFormat sdFormat = getSimpleDateFormat(pattern);
        Date date;
        try {
            value = formatDate(value, pattern);
            date = sdFormat.parse(value);
        } catch (Exception e) {
            throw new RuntimeException("时间格式转换失败");
        }
        return date;
    }

    /**
     * 当前日期格式化 yyyy-MM-dd
     *
     * @return yyyy-MM-dd
     */
    public static String format() {
        DateFormat format = getSimpleDateFormat("yyyy-MM-dd");
        return format.format(new Date());
    }

    /**
     * 当前时间格式化
     *
     * @param pattern 时间格式
     * @return 格式化后的时间字符串
     */
    public static String format(String pattern) {
        SimpleDateFormat format = getSimpleDateFormat(pattern);
        return format.format(new Date());
    }

    /**
     * 当前时间格式化
     *
     * @param date    被格式化的时间
     * @param pattern 时间格式
     * @return 格式化后的时间字符串
     */
    public static String format(Date date, String pattern) {
        SimpleDateFormat format = getSimpleDateFormat(pattern);
        return format.format(date);
    }

    /**
     * 字符串时间格式化
     *
     * @param dateStr 时间字符串
     * @param pattern 时间格式
     * @return 格式化后的时间
     */
    public static String formatDate(String dateStr, String pattern) {
        if (dateStr == null || "".equals(dateStr)) {
            return "";
        }
        Date date;
        DateFormat formatIn;
        DateFormat formatOut;
        ParsePosition pos = new ParsePosition(0);
        dateStr = dateStr.replace("-", "").replace(":", "");
        if ("".equals(dateStr.trim())) {
            return "";
        }
        try {
            if (Long.parseLong(dateStr) == 0L) {
                return "";
            }
        } catch (Exception e) {
            return dateStr;
        }
        try {
            switch (dateStr.trim().length()) {
                case 14:
                    formatIn = getSimpleDateFormat("yyyyMMddHHmmss");
                    break;
                case 12:
                    formatIn = getSimpleDateFormat("yyyyMMddHHmm");
                    break;
                case 10:
                    formatIn = getSimpleDateFormat("yyyyMMddHH");
                    break;
                case 8:
                    formatIn = getSimpleDateFormat("yyyyMMdd");
                    break;
                case 6:
                    formatIn = getSimpleDateFormat("yyyyMM");
                    break;
                case 7:
                case 9:
                case 11:
                case 13:
                default:
                    return dateStr;
            }
            date = formatIn.parse(dateStr, pos);
            if (date == null) {
                return dateStr;
            }
            if ((pattern == null) || ("".equals(pattern.trim()))) {
                formatOut = getSimpleDateFormat("yyyy年MM月dd日");
            } else {
                formatOut = getSimpleDateFormat(pattern);
            }
            return formatOut.format(date);
        } catch (Exception ex) {
            return dateStr;
        }
    }
}


















