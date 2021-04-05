package com.github.helloichen.excel;

import com.github.helloichen.annotation.IChenExcelField;
import com.github.helloichen.util.DateUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author iChen
 * @date 2021-04-05
 */
@Slf4j
public class ExcelReader {

    /**
     * 解析Excel 支持2003、2007<br>
     * 利用反射技术完成ExcelField注解解析
     * properties、obj需要符合如下规则：<br>
     *
     * @param file 待解析的Excel文件
     * @param obj  反射对象的Class
     */
    public static <T> List<T> read(File file, Class<T> obj) throws Exception {
        Workbook book;
        try {
            //解析2003
            book = new HSSFWorkbook(new FileInputStream(file));
        } catch (Exception e) {
            //解析2007
            book = new XSSFWorkbook(new FileInputStream(file));
        }
        return getExcelContent(book, obj);
    }

    public static <T> List<T> read(InputStream file, Class<T> obj) throws Exception {
        Workbook book;
        try {
            //解析2003
            book = new XSSFWorkbook(file);
        } catch (Exception e) {
            //解析2007
            book = new HSSFWorkbook(file);
        }
        return getExcelContent(book, obj);
    }

    /**
     * 解析Excel 支持2003、2007<br>
     * 利用反射技术完成ExcelField注解解析
     *
     * @param filePath 待解析的Excel文件的路径
     * @param obj      反射对象的Class
     */
    public static <T> List<T> read(String filePath, Class<T> obj) throws Exception {
        File file = new File(filePath);
        if (!file.exists()) {
            throw new Exception("指定的文件不存在");
        }
        return read(file, obj);
    }

    /**
     * 根据params、object解析Excel，并且构建list集合
     *
     * @param book  WorkBook对象，他代表了待将解析的Excel文件
     * @param clazz 构建的Object对象，每一个row都相当于一个object对象
     */
    private static <T> List<T> getExcelContent(Workbook book,
                                               Class<T> clazz) throws Exception {
        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        Map<String, Method> methodTempMap = Arrays.stream(methods)
                .filter(method -> Modifier.isPublic(method.getModifiers()))
                .filter(method -> method.getName().startsWith("set"))
                .collect(Collectors.toMap(Method::getName, Function.identity()));

        Map<String, Method> setterMethodMap = new HashMap<>(16);
        List<Field> fieldList = new ArrayList<>();
        for (Field field : fields) {
            IChenExcelField IChenExcelField = field.getAnnotation(IChenExcelField.class);
            if (IChenExcelField != null && IChenExcelField.exportField()) {
                String attr = field.getName();
                String setterName = "set" + Character.toUpperCase(attr.charAt(0)) + attr.substring(1);
                Method method = methodTempMap.get(setterName);
                if (method == null) {
                    // setter方法不为null才可以导出
                    log.warn("{}类{}属性无setter方法，无法导导入", clazz.getName(), attr);
                    continue;
                }
                fieldList.add(field);
                setterMethodMap.put(attr, method);
            }
        }

        //初始化结果集
        List<T> resultList = new ArrayList<>();
        for (int numSheet = 0; numSheet < book.getNumberOfSheets(); numSheet++) {
            Sheet sheet = book.getSheetAt(numSheet);
            //谨防中间空一行
            if (sheet == null) {
                continue;
            }

            //一个row就相当于一个Object，标题行不读取
            for (int numRow = 1; numRow <= sheet.getLastRowNum(); numRow++) {
                Row row = sheet.getRow(numRow);
                if (isBlankRow(row)) {
                    continue;
                }
                resultList.add(getObject(row, setterMethodMap, fieldList, clazz));
            }
        }
        return resultList;
    }

    private static boolean isBlankRow(Row row) {
        if (row == null) {
            return true;
        }
        Cell cell = row.getCell(0, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
        return cell == null || cell.getCellTypeEnum().equals(CellType.BLANK);
    }

    /**
     * 获取row的数据，利用反射机制构建Object对象
     *
     * @param row       row对象
     * @param methodMap object对象的setter方法映射
     * @param fieldList object对象的属性
     */
    private static <T> T getObject(Row row,
                                   Map<String, Method> methodMap, List<Field> fieldList, Class<T> obj) throws Exception {
        T object = obj.newInstance();
        for (int numCell = 0; numCell < fieldList.size(); numCell++) {
            Field field = fieldList.get(numCell);
            Cell cell = row.getCell(numCell);
            if (cell == null) {
                continue;
            }

            cell.setCellType(CellType.STRING);

            String cellValue = getValue(cell);
            //在object对象中对应的setter方法
            Method method = methodMap.get(field.getName());
            setObjectPropertyValue(object, field, method, cellValue);
        }
        return object;
    }

    /**
     * 根据指定属性的的setter方法给object对象设置值
     *
     * @param obj    object对象
     * @param field  object对象的属性
     * @param method object对象属性的相对应的方法
     * @param value  需要设置的值
     */
    private static void setObjectPropertyValue(Object obj, Field field,
                                               Method method, String value) throws Exception {
        Object[] oo = new Object[1];

        String type = field.getType().getName();
        switch (type) {
            case "java.lang.String":
            case "String":
                oo[0] = value;
                break;
            case "java.lang.Integer":
            case "java.lang.int":
            case "Integer":
            case "int":
                if (value.length() > 0) {
                    oo[0] = Integer.valueOf(value);
                }
                break;
            case "java.lang.Float":
            case "java.lang.float":
            case "Float":
            case "float":
                if (value.length() > 0) {
                    oo[0] = Float.valueOf(value);
                }
                break;
            case "java.lang.Double":
            case "java.lang.double":
            case "Double":
            case "double":
                if (value.length() > 0) {
                    oo[0] = Double.valueOf(value);
                }
                break;
            case "java.math.BigDecimal":
            case "BigDecimal":
                if (value.length() > 0) {
                    oo[0] = new BigDecimal(value);
                }
                break;
            case "java.util.Date":
            case "Date":
                //当长度为19(yyyy-MM-dd HH24:mm:ss)或者为14(yyyyMMddHH24mmss)时Date格式转换为yyyyMMddHH24mmss
                if (value.length() > 0) {
                    if (value.length() == 19 || value.length() == 14) {
                        oo[0] = DateUtils.string2Date(value, "yyyyMMddHH24mmss");
                    } else {
                        //其余全部转换为yyyyMMdd格式
                        oo[0] = DateUtils.string2Date(value, "yyyyMMdd");
                    }
                }
                break;
            case "java.sql.Timestamp":
                if (value.length() > 0) {
                    oo[0] = DateUtils.formatDate(value, "yyyyMMddHH24mmss");
                }
                break;
            case "java.lang.Boolean":
            case "Boolean":
                if (value.length() > 0) {
                    oo[0] = Boolean.valueOf(value);
                }
                break;
            case "java.lang.Long":
            case "java.lang.long":
            case "Long":
            case "long":
                if (value.length() > 0) {
                    oo[0] = Long.valueOf(value);
                }
                break;
            default:
        }
        method.invoke(obj, oo);
    }

    private static String getValue(Cell cell) {
        if (CellType.BOOLEAN.equals(cell.getCellTypeEnum())) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (CellType.NUMERIC.equals(cell.getCellTypeEnum())) {
            return NumberToTextConverter.toText(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }

}
