package com.github.andyczy.java.excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.OutputStream;
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
public class LocalExcelUtils {


    private static Logger log = LoggerFactory.getLogger(LocalExcelUtils.class);
    private static final ThreadLocal<SimpleDateFormat> fmt = new ThreadLocal<>();
    private static final ThreadLocal<DecimalFormat> df = new ThreadLocal<>();
    private static final ThreadLocal<LocalExcelUtils> UTILS_THREAD_LOCAL = new ThreadLocal<>();
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

    public static final LocalExcelUtils initialization() {
        LocalExcelUtils excelUtils = UTILS_THREAD_LOCAL.get();
        if (excelUtils == null) {
            excelUtils = new LocalExcelUtils();
            UTILS_THREAD_LOCAL.set(excelUtils);
        }
        return excelUtils;
    }

    public LocalExcelUtils() {
        filePath = this.getFilePath();
        dataLists = this.getDataLists();
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
        defaultColumnWidth = this.getDefaultColumnWidth();
        fontSize = this.getFontSize();
    }


    /**
     * 本地输出
     * Excel导出：无样式（行、列、单元格样式）、自适应列宽
     *
     * @return
     */
    public Boolean localNoStyleNoResponse() {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        SXSSFRow sxssfRow = null;
        try {
            setDataListNoStyle(sxssfWorkbook, sxssfRow, dataLists, regionMap, mapColumnWidth, paneMap, sheetName, labelName, dropDownMap, defaultColumnWidth,fontSize);
            setIo(sxssfWorkbook, filePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }

    /**
     * 本地输出
     * Excel导出：有样式（行、列、单元格样式）、自适应列宽
     *
     * @return
     */
    public Boolean localNoResponse() {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        SXSSFRow sxssfRow = null;
        try {
            setDataList(sxssfWorkbook, sxssfRow, dataLists, regionMap, mapColumnWidth, styles, paneMap, sheetName, labelName, rowStyles, columnStyles, dropDownMap, defaultColumnWidth,fontSize);
            setIo(sxssfWorkbook, filePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
        log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
        return true;
    }


    /**
     * 功能描述:
     * 1.excel 模板数据导入。
     * <p>
     * 更新日志:
     * 1.共用获取Excel表格数据。
     * 2.多单元数据获取。
     * 3.多单元从第几行开始获取数据[2018-09-20]
     * 4.多单元根据那些列为空来忽略行数据[2018-10-22]
     * <p>
     * 版  本:
     * 1.apache poi 3.17
     * 2.apache poi-ooxml  3.17
     *
     * @param book           Workbook对象（不可为空）
     * @param sheetName      多单元数据获取（不可为空）
     * @param indexMap       多单元从第几行开始获取数据，默认从第二行开始获取（可为空，如 hashMapIndex.put(1,3); 第一个表格从第三行开始获取）
     * @param continueRowMap 多单元根据那些列为空来忽略行数据（可为空，如 mapContinueRow.put(1,new Integer[]{1, 3}); 第一个表格从1、3列为空就忽略）
     * @return
     */
    @SuppressWarnings({"deprecation", "rawtypes"})
    public static List<List<LinkedHashMap<String, String>>> importForExcelData(Workbook book, String[] sheetName, HashMap indexMap, HashMap continueRowMap) {
        long startTime = System.currentTimeMillis();
        log.info("=== ===  === :Excel tool class export start run!");
        try {
            List<List<LinkedHashMap<String, String>>> returnDataList = new ArrayList<>();
            for (int k = 0; k <= sheetName.length - 1; k++) {
                //  得到第K个工作表对象、得到第K个工作表中的总行数。
                Sheet sheet = book.getSheetAt(k);
                int rowCount = sheet.getLastRowNum() + 1;
                Row valueRow = null;

                List<LinkedHashMap<String, String>> rowListValue = new ArrayList<>();
                LinkedHashMap<String, String> cellHashMap = null;

                int irow = 1;
                //  第n个工作表:从开始获取数据、默认第一行开始获取。
                if (indexMap != null && indexMap.get(k + 1) != null) {
                    irow = Integer.valueOf(indexMap.get(k + 1).toString()) - 1;
                }
                //  第n个工作表:数据获取。
                for (int i = irow; i < rowCount; i++) {
                    valueRow = sheet.getRow(i);
                    if (valueRow == null) {
                        continue;
                    }

                    //  第n个工作表:从开始列忽略数据、为空就跳过。
                    if (continueRowMap != null && continueRowMap.get(k + 1) != null) {
                        int continueRowCount = 0;
                        Integer[] continueRow = (Integer[]) continueRowMap.get(k + 1);
                        for (int w = 0; w <= continueRow.length - 1; w++) {
                            Cell valueRowCell = valueRow.getCell(continueRow[w] - 1);
                            if (valueRowCell == null || isBlank(valueRowCell.toString())) {
                                continueRowCount = continueRowCount + 1;
                            }
                        }
                        if (continueRowCount == continueRow.length) {
                            continue;
                        }
                    }

                    cellHashMap = new LinkedHashMap<>();

                    //  第n个工作表:获取列数据。
                    for (int j = 0; j < valueRow.getLastCellNum(); j++) {
                        cellHashMap.put(Integer.toString(j), getCellVal(valueRow.getCell(j)));
                    }
                    if (cellHashMap.size() > 0) {
                        rowListValue.add(cellHashMap);
                    }
                }
                returnDataList.add(rowListValue);
            }
            log.info("=== ===  === :Excel tool class export run time:" + (System.currentTimeMillis() - startTime) + " ms!");
            return returnDataList;
        } catch (Exception e) {
            log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：Excel tool class export exception !");
            e.printStackTrace();
            return null;
        }
    }


    /**
     * 输出本地地址
     *
     * @param sxssfWorkbook
     * @param filePath
     * @throws Exception
     */
    private static void setIo(SXSSFWorkbook sxssfWorkbook, String filePath) throws Exception {
        OutputStream outputStream = null;
        try {
            if (filePath != null) {
                outputStream = new FileOutputStream(filePath);
                if (outputStream != null) {
                    sxssfWorkbook.write(outputStream);
                    sxssfWorkbook.dispose();
                    outputStream.flush();
                    outputStream.close();
                }
            }
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
                        val = LocalExcelUtils.initialization().getDateFormat().format(cell.getDateCellValue());
                    } else {
                        val = LocalExcelUtils.initialization().getDecimalFormat().format(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    if (cell.getStringCellValue().trim().length() >= DATE_LENGTH && verificationDate(cell.getStringCellValue(), LocalExcelUtils.initialization().dateFormatStr)) {
                        val = strToDateFormat(cell.getStringCellValue(), LocalExcelUtils.initialization().dateFormatStr, LocalExcelUtils.initialization().expectDateFormatStr);
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
    /**
     * 默认列宽大小：默认16
     */
    private Integer defaultColumnWidth;
    /**
     * 默认字体大小：默认12号字体
     */
    private Integer fontSize;

    public Integer getFontSize() {
        if (fontSize == null) {
            fontSize = 12;
        }
        return fontSize;
    }

    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    public Integer getDefaultColumnWidth() {
        if (defaultColumnWidth == null) {
            defaultColumnWidth = 16;
        }
        return defaultColumnWidth;
    }

    public void setDefaultColumnWidth(Integer defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
    }

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

