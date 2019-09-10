package com.github.andyczy.java.excel;


import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFPicture;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static com.github.andyczy.java.excel.CommonsUtils.*;


/**
 * @author Chenzy
 * @date 2018-05-03
 * @description: Excel 导入导出操作相关工具类
 * @Email: 649954910@qq.com
 */
public class ExcelUtils {


    private static Logger log = LoggerFactory.getLogger(ExcelUtils.class);
    private static final ThreadLocal<SimpleDateFormat> fmt = new ThreadLocal<>();
    private static final ThreadLocal<DecimalFormat> df = new ThreadLocal<>();
    private static final ThreadLocal<ExcelUtils> UTILS_THREAD_LOCAL = new ThreadLocal<>();
    private static final Integer DATE_LENGTH = 10;

    //    private static final String MESSAGE_FORMAT_df = "#.######";
    //    private static final String MESSAGE_FORMAT = "yyyy-MM-dd";


    private SimpleDateFormat getDateFormat() {
        SimpleDateFormat format = fmt.get();
        if (format == null) {
            //默认格式日期： "yyyy-MM-dd"
            format = new SimpleDateFormat(expectDateFormatStr, Locale.getDefault());
            fmt.set(format);
        }
        return format;
    }

    public DecimalFormat getDecimalFormat() {
        DecimalFormat format = df.get();
        if (format == null) {
            //默认数字格式： "#.######" 六位小数点
            format = new DecimalFormat(numeralFormat);
            df.set(format);
        }
        return format;
    }

    public static final ExcelUtils initialization() {
        ExcelUtils excelUtils = UTILS_THREAD_LOCAL.get();
        if (excelUtils == null) {
            excelUtils = new ExcelUtils();
            UTILS_THREAD_LOCAL.set(excelUtils);
        }
        return excelUtils;
    }

    public ExcelUtils() {
        filePath = this.getFilePath();
        dataLists = this.getDataLists();
        response = this.getResponse();
        regionMap = this.getRegionMap();
        mapColumnWidth = this.getMapColumnWidth();
        styles = this.getStyles();
        paneMap = this.getPaneMap();
        fileName = this.getFileName();
        sheetName = this.getSheetName();
        labelName = this.getLabelName();
        rowStyles = this.getRowStyles();
        columnStyles = this.getColumnStyles();
        dropDownMap = this.getDropDownMap();
        numeralFormat = this.getNumeralFormat();
        dateFormatStr = this.getDateFormatStr();
        expectDateFormatStr = this.getExpectDateFormatStr();
    }


    /**
     * web 响应（response）
     * Excel导出：有样式（行、列、单元格样式）、自适应列宽
     * 功能描述: excel 数据导出、导出模板
     * 更新日志:
     * 1.response.reset();注释掉reset，否在会出现跨域错误。[2018-05-18]
     * 2.新增导出多个单元。[2018-08-08]
     * 3.poi官方建议大数据量解决方案：SXSSFWorkbook。[2018-08-08]
     * 4.自定义下拉列表：对每个单元格自定义下拉列表。[2018-08-08]
     * 5.数据遍历方式换成数组(效率较高)。[2018-08-08]
     * 6.可提供模板下载。[2018-08-08]
     * 7.每个表格的大标题[2018-09-14]
     * 8.自定义列宽：对每个单元格自定义列宽[2018-09-18]
     * 9.自定义样式：对每个单元格自定义样式[2018-10-22]
     * 10.自定义单元格合并：对每个单元格合并[2018-10-22]
     * 11.固定表头[2018-10-23]
     * 12.自定义样式：单元格自定义某一列或者某一行样式[2018-11-12]
     * 13.忽略边框(默认是有边框)[2018-11-15]
     * 14.函数式编程换成面向对象编程[2018-12-06-5]
     * 15.单表百万数据量导出时样式设置过多，导致速度慢（行、列、单元格样式暂时控制10万行、超过无样式）[2019-01-30]
     * 版  本:
     * 1.apache poi 3.17
     * 2.apache poi-ooxml  3.17
     *
     * @return
     */
    public Boolean exportForExcelsOptimize() {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            // 设置数据
            setDataList(sxssfWorkbook, sxssfRow, dataLists, regionMap, mapColumnWidth, styles, paneMap, sheetName, labelName, rowStyles, columnStyles, dropDownMap);
            // io 响应
            setIo(sxssfWorkbook, outputStream, fileName, sheetName, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }


    /**
     * Excel导出：无样式（行、列、单元格样式）、自适应列宽
     * web 响应（response）
     *
     * @return
     */
    public Boolean exportForExcelsNoStyle() {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            setDataListNoStyle(sxssfWorkbook, sxssfRow, dataLists, regionMap, mapColumnWidth, paneMap, sheetName, labelName, dropDownMap);
            setIo(sxssfWorkbook, outputStream, fileName, sheetName, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }


    /**
     * 功能描述: excel 数据导出、导出模板
     * <p>
     * 更新日志:
     * 1.response.reset();注释掉reset，否在会出现跨域错误。[2018-05-18]
     * 2.新增导出多个单元。[2018-08-08]
     * 3.poi官方建议大数据量解决方案：SXSSFWorkbook。[2018-08-08]
     * 4.自定义下拉列表：对每个单元格自定义下拉列表。[2018-08-08]
     * 5.数据遍历方式换成数组(效率较高)。[2018-08-08]
     * 6.可提供模板下载。[2018-08-08]
     * 7.每个表格的大标题[2018-09-14]
     * 8.自定义列宽：对每个单元格自定义列宽[2018-09-18]
     * 9.自定义样式：对每个单元格自定义样式[2018-10-22]
     * 10.自定义单元格合并：对每个单元格合并[2018-10-22]
     * 11.固定表头[2018-10-23]
     * 12.自定义样式：单元格自定义某一列或者某一行样式[2018-11-12]
     * 13.忽略边框(默认是有边框)[2018-11-15]
     * 14.函数式编程换成面向对象编程[2018-12-06-5]
     * 15.单表百万数据量导出时样式设置过多，导致速度慢（行、列、单元格样式暂时去掉）[2019-01-30]
     * <p>
     * 版  本:
     * 1.apache poi 3.17
     * 2.apache poi-ooxml  3.17
     *
     * @param response
     * @param dataLists    导出的数据(不可为空：如果只有标题就导出模板)
     * @param sheetName    sheet名称（不可为空）
     * @param columnMap    自定义：对每个单元格自定义列宽（可为空）
     * @param dropDownMap  自定义：对每个单元格自定义下拉列表（可为空）
     * @param styles       自定义：每一个单元格样式（可为空）
     * @param rowStyles    自定义：某一行样式（可为空）
     * @param columnStyles 自定义：某一列样式（可为空）
     * @param regionMap    自定义：单元格合并（可为空）
     * @param paneMap      固定表头（可为空）
     * @param labelName    每个表格的大标题（可为空）
     * @param fileName     文件名称(可为空，默认是：sheet 第一个名称)
     * @param notBorderMap 忽略边框(默认是有边框)
     * @return
     */
    @Deprecated
    public static Boolean exportForExcel(HttpServletResponse response, List<List<String[]>> dataLists, HashMap notBorderMap,
                                         HashMap regionMap, HashMap columnMap, HashMap styles, HashMap paneMap, String fileName,
                                         String[] sheetName, String[] labelName, HashMap rowStyles, HashMap columnStyles, HashMap dropDownMap) {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            // 设置数据
            setDataList(sxssfWorkbook, sxssfRow, dataLists, regionMap, columnMap, styles, paneMap, sheetName, labelName, rowStyles, columnStyles, dropDownMap);
            // io 响应
            setIo(sxssfWorkbook, outputStream, fileName, sheetName, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }


    /**
     * response 响应
     *
     * @param sxssfWorkbook
     * @param outputStream
     * @param fileName
     * @param sheetName
     * @param response
     * @throws Exception
     */
    private static void setIo(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream, String fileName, String[] sheetName, HttpServletResponse response) throws Exception {
        try {
            if (response != null) {
                response.setHeader("Charset", "UTF-8");
                response.setHeader("Content-Type", "application/force-download");
                response.setHeader("Content-Type", "application/vnd.ms-excel");
                response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(fileName == null ? sheetName[0] : fileName, "utf8") + ".xlsx");
                response.flushBuffer();
                outputStream = response.getOutputStream();
            }
            writeAndColse(sxssfWorkbook, outputStream);
        } catch (Exception e) {
            e.getSuppressed();
        }
    }




    /**
     * 功能描述: 获取Excel单元格中的值并且转换java类型格式
     *
     * @param cell
     * @return
     */
    private static String getCellVal(Cell cell) {
        String val = null;
        if (cell != null) {
            CellType cellType = cell.getCellType();
            switch (cellType) {
                case NUMERIC:
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                        val = ExcelUtils.initialization().getDateFormat().format(cell.getDateCellValue());
                    } else {
                        val = ExcelUtils.initialization().getDecimalFormat().format(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    if (cell.getStringCellValue().trim().length() >= DATE_LENGTH && verificationDate(cell.getStringCellValue(), ExcelUtils.initialization().dateFormatStr)) {
                        val = strToDateFormat(cell.getStringCellValue(), ExcelUtils.initialization().dateFormatStr, ExcelUtils.initialization().expectDateFormatStr);
                    } else {
                        val = cell.getStringCellValue();
                    }
                    break;
                case BOOLEAN:
                    val = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK:
                    val = cell.getStringCellValue();
                    break;
                case ERROR:
                    val = "错误";
                    break;
                case FORMULA:
                    try {
                        val = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        val = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                default:
                    val = cell.getRichStringCellValue() == null ? null : cell.getRichStringCellValue().toString();
            }
        } else {
            val = "";
        }
        return val;
    }

    /**
     * 导出数据必填
     */
    private List<List<String[]>> dataLists;
    /**
     * sheet名称必填
     */
    private String[] sheetName;
    /**
     * 每个表格的大标题
     */
    private String[] labelName;
    /**
     * 页面响应
     */
    private HttpServletResponse response;

    /**
     * 自定义：单元格合并
     */
    private HashMap regionMap;
    /**
     * 自定义：对每个单元格自定义列宽
     */
    private HashMap mapColumnWidth;
    /**
     * 自定义：每一个单元格样式
     */
    private HashMap styles;
    /**
     * 自定义：固定表头
     */
    private HashMap paneMap;
    /**
     * 自定义：某一行样式
     */
    private HashMap rowStyles;
    /**
     * 自定义：某一列样式
     */
    private HashMap columnStyles;
    /**
     * 自定义：对每个单元格自定义下拉列表
     */
    private HashMap dropDownMap;
    /**
     * 文件名称
     */
    private String fileName;
    /**
     * 导出本地路径
     */
    private String filePath;

    /**
     * 导出数字格式化：默认是保留六位小数点
     */
    private String numeralFormat;


    /**
     * 导出日期格式化：默认是"yyyy-MM-dd"格式
     */
    private String dateFormatStr;
    /**
     * 期望转换后的日期格式：默认是 dateFormatStr
     */
    private String expectDateFormatStr;


    public void setDateFormatStr(String dateFormatStr) {
        if (dateFormatStr == null) {
            dateFormatStr = "yyyy-MM-dd";
        }
        this.dateFormatStr = dateFormatStr;
    }

    public String getDateFormatStr() {
        if (dateFormatStr == null) {
            dateFormatStr = "yyyy-MM-dd";
        }
        return dateFormatStr;
    }

    public String getExpectDateFormatStr() {
        if (expectDateFormatStr == null) {
            expectDateFormatStr = dateFormatStr;
        }
        return expectDateFormatStr;
    }

    public void setExpectDateFormatStr(String expectDateFormatStr) {
        if (expectDateFormatStr == null) {
            expectDateFormatStr = dateFormatStr;
        }
        this.expectDateFormatStr = expectDateFormatStr;
    }

    public void setNumeralFormat(String numeralFormat) {
        if (numeralFormat == null) {
            numeralFormat = "#.######";
        }
        this.numeralFormat = numeralFormat;
    }

    public String getNumeralFormat() {
        if (numeralFormat == null) {
            numeralFormat = "#.######";
        }
        return numeralFormat;
    }


    public List<List<String[]>> getDataLists() {
        return dataLists;
    }

    public void setDataLists(List<List<String[]>> dataLists) {
        this.dataLists = dataLists;
    }

    public String[] getSheetName() {
        return sheetName;
    }

    public void setSheetName(String[] sheetName) {
        this.sheetName = sheetName;
    }

    public String[] getLabelName() {
        return labelName;
    }

    public void setLabelName(String[] labelName) {
        this.labelName = labelName;
    }

    public HttpServletResponse getResponse() {
        return response;
    }

    public void setResponse(HttpServletResponse response) {
        this.response = response;
    }


    public HashMap getRegionMap() {
        return regionMap;
    }

    public void setRegionMap(HashMap regionMap) {
        this.regionMap = regionMap;
    }

    public HashMap getMapColumnWidth() {
        return mapColumnWidth;
    }

    public void setMapColumnWidth(HashMap mapColumnWidth) {
        this.mapColumnWidth = mapColumnWidth;
    }

    public HashMap getStyles() {
        return styles;
    }

    public void setStyles(HashMap styles) {
        this.styles = styles;
    }

    public HashMap getPaneMap() {
        return paneMap;
    }

    public void setPaneMap(HashMap paneMap) {
        this.paneMap = paneMap;
    }

    public HashMap getRowStyles() {
        return rowStyles;
    }

    public void setRowStyles(HashMap rowStyles) {
        this.rowStyles = rowStyles;
    }

    public HashMap getColumnStyles() {
        return columnStyles;
    }

    public void setColumnStyles(HashMap columnStyles) {
        this.columnStyles = columnStyles;
    }

    public HashMap getDropDownMap() {
        return dropDownMap;
    }

    public void setDropDownMap(HashMap dropDownMap) {
        this.dropDownMap = dropDownMap;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

}

