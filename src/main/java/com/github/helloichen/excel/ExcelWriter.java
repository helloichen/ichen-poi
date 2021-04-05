package com.github.helloichen.excel;

import com.alibaba.fastjson.JSON;
import com.github.helloichen.annotation.IChenExcelField;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.net.URLEncoder;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author iChen
 * @date 2021-04-05
 */
@Slf4j
public class ExcelWriter {
    /**
     * 使用浏览器选择路径下载
     *
     * @param response 相应对象
     * @param fileName 导出文件名称
     * @param rowData  excel数据
     */
    public static void write(HttpServletResponse response, String fileName,
                             List<?> rowData, Class<?> clazz) {

        response.setContentType("application/binary;charset=UTF-8");
        // 进行转码，使其支持中文文件名
        try {
            fileName = URLEncoder.encode(fileName, "UTF-8");
        } catch (UnsupportedEncodingException e1) {
            log.error("出现转码异常");
        }
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
        writeExcelFile(rowData, clazz, response);
    }

    /**
     * 写excel文件
     *
     * @param rowData  表数据list
     * @param clazz    导出数据所属类型
     * @param response 响应对象
     */
    private static void writeExcelFile(List<?> rowData, Class<?> clazz, HttpServletResponse response) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        try (OutputStream outputStream = response.getOutputStream()) {
            XSSFSheet sheet = workbook.createSheet();
            try {
                writeExcelContent(workbook, sheet, rowData, clazz);
                workbook.write(outputStream);
            } catch (Exception e) {
                // 重置response
                log.error("导出文件失败,错误信息：{}", e.getMessage());
                response.reset();
                response.setContentType("application/json");
                response.setCharacterEncoding("utf-8");
                Map<String, String> map = new HashMap<>(8);
                map.put("code", "-1");
                map.put("data", "下载文件失败");
                map.put("message", "操作失败");
                try (PrintWriter writer = response.getWriter()) {
                    writer.println(JSON.toJSONString(map));
                } catch (Exception ioExp) {
                    log.error("失败信息返回出错！");
                }
            }
        } catch (Exception e) {
            log.error("导出文件失败,错误信息：{}", e.getMessage());
        }
    }

    /**
     * 写excel内容
     *
     * @param workbook 工作簿
     * @param sheet    工作表
     * @param rowData  表数据list
     * @param clazz    导出数据所属类型
     */
    private static void writeExcelContent(XSSFWorkbook workbook, XSSFSheet sheet, List<?> rowData, Class<?> clazz) throws Exception {

        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        Map<String, Method> methodTempMap = Arrays.stream(methods)
                .filter(method -> Modifier.isPublic(method.getModifiers()))
                .filter(method -> method.getName().startsWith("get"))
                .collect(Collectors.toMap(Method::getName, Function.identity()));

        Map<String, Method> getterMethodMap = new HashMap<>(16);
        Map<String, String> titleMap = new HashMap<>(16);
        List<Field> fieldList = new ArrayList<>();
        for (Field field : fields) {
            IChenExcelField excelField = field.getAnnotation(IChenExcelField.class);
            if (excelField != null && excelField.importField()) {
                String attr = field.getName();
                String getterName = "get" + Character.toUpperCase(attr.charAt(0)) + attr.substring(1);
                Method method = methodTempMap.get(getterName);
                // getter方法不为null才可以导出
                if (method == null) {
                    log.warn("{}类{}属性无getter方法，无法导出", clazz.getName(), attr);
                    continue;
                }
                fieldList.add(field);
                String title = excelField.value();
                titleMap.put(attr, title);
                getterMethodMap.put(attr, method);
            }
        }

        //表头
        writeTitlesToExcel(workbook, sheet, fieldList, titleMap);
        //数据
        writeRowDataToExcel(workbook, sheet, rowData, fieldList, getterMethodMap);
    }

    /**
     * 写入工作表数据
     *
     * @param workbook  工作簿
     * @param sheet     工作表
     * @param rowData   工作表数据
     * @param methodMap 属性Getter方法Map
     * @throws Exception 抛出异常
     */
    private static void writeRowDataToExcel(XSSFWorkbook workbook, XSSFSheet sheet, List<?> rowData,
                                            List<Field> fieldList, Map<String, Method> methodMap) throws Exception {
        Font dataFont = workbook.createFont();
        dataFont.setFontName("simsun");
        dataFont.setFontHeightInPoints((short) 14);
        dataFont.setColor(IndexedColors.BLACK.index);
        XSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setFont(dataFont);
        setBorder(dataStyle, new XSSFColor(new java.awt.Color(0, 0, 0)));

        int colIndex;
        //从第一行开始写数据，第0行为标题行
        int rowIndex = 1;
        for (Object data : rowData) {
            Row dataRow = sheet.createRow(rowIndex);
            dataRow.setHeightInPoints(25);
            colIndex = 0;
            for (Field field : fieldList) {
                String fieldName = field.getName();
                Cell cell = dataRow.createCell(colIndex);
                Method method = methodMap.get(fieldName);
                Object invoke = method.invoke(data);
                if (invoke != null) {
                    cell.setCellValue(invoke.toString());
                } else {
                    cell.setCellValue("");
                }
                cell.setCellStyle(dataStyle);
                colIndex++;
            }
            rowIndex++;
        }

    }

    /**
     * 设置工作表表头
     *
     * @param workbook excel工作簿
     * @param sheet    excel工作表
     */
    private static void writeTitlesToExcel(XSSFWorkbook workbook, XSSFSheet sheet, List<Field> fieldList, Map<String, String> titleMap) {
        Font titleFont = workbook.createFont();
        //设置字体
        titleFont.setFontName("黑体");
        //设置粗体
        titleFont.setBold(true);
        //设置字号
        titleFont.setFontHeightInPoints((short) 14);
        //设置颜色
        titleFont.setColor(IndexedColors.BLACK.index);
        XSSFCellStyle titleStyle = workbook.createCellStyle();
        //水平居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置图案颜色
        titleStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 255, 0)));
        //设置图案样式
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        setBorder(titleStyle, new XSSFColor(new java.awt.Color(0, 0, 0)));

        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(25);

        int colIndex = 0;
        for (Field field : fieldList) {
            Cell cell = titleRow.createCell(colIndex);
            cell.setCellValue(titleMap.get(field.getName()));
            cell.setCellStyle(titleStyle);
            colIndex++;
        }
    }

    /**
     * 设置边框
     *
     * @param style XSSFCellStyle
     * @param color XSSFColor
     */

    private static void setBorder(XSSFCellStyle style, XSSFColor color) {

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, color);

    }

    /**
     * 自动调整列宽
     *
     * @param sheet        sheet工作表
     * @param columnNumber 列数
     */

    private static void autoSizeColumns(Sheet sheet, int columnNumber) {
        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);
            sheet.autoSizeColumn(i, true);
            int newWidth = sheet.getColumnWidth(i) + 100;
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, Math.min(newWidth, 65280));
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }
}
