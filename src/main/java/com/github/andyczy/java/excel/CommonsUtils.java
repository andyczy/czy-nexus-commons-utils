package com.github.andyczy.java.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.util.CellUtil.createCell;

/**
 * Author: Chen zheng you
 * CreateTime: 2019-02-15 16:10
 * Description:
 */
public class CommonsUtils {

    private static Logger log = LoggerFactory.getLogger(CommonsUtils.class);

    private static final String DataValidationError1 = "Excel表格提醒：";
    private static final String DataValidationError2 = "数据不规范，请选择表格下拉列表中的数据！";
    public static final Integer MAX_ROWSUM = 1048570;
    public static final Integer MAX_ROWSYTLE = 100000;


    /**
     * 设置数据：有样式（行、列、单元格样式）
     *
     * @param sxssfRow
     * @param dataLists
     * @param notBorderMap
     * @param regionMap
     * @param columnMap
     * @param styles
     * @param paneMap
     * @param sheetName
     * @param labelName
     * @param rowStyles
     * @param columnStyles
     * @param dropDownMap
     */
    public static void setDataList(SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<String[]>> dataLists, HashMap notBorderMap,
                                   HashMap regionMap, HashMap columnMap, HashMap styles, HashMap paneMap,
                                   String[] sheetName, String[] labelName, HashMap rowStyles, HashMap columnStyles, HashMap dropDownMap) throws Exception {
        if (dataLists == null) {
            log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (sheetName == null) {
            log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        for (List<String[]> listRow : dataLists) {
            SXSSFSheet sxssfSheet = wb.createSheet();
            sxssfSheet.setDefaultColumnWidth((short) 16);
            wb.setSheetName(k, sheetName[k]);
            CellStyle cellStyle = wb.createCellStyle();
            XSSFFont font = (XSSFFont) wb.createFont();
            int jRow = 0;
            if (labelName != null) {
                //  自定义：大标题和样式。参数说明：new String[]{"表格数据一", "表格数据二", "表格数据三"}
                sxssfRow = sxssfSheet.createRow(0);
                Cell cell = createCell(sxssfRow, 0, labelName[k]);
                setMergedRegion(sxssfSheet, 0, 0, 0, listRow.get(0).length - 1);
                setLabelStyles(wb, cell, sxssfRow);
                jRow = 1;
            }
            //  自定义：每个表格固定表头（看该方法说明）。
            Integer pane = 1;
            if (paneMap != null && paneMap.get(k + 1) != null) {
                pane = (Integer) paneMap.get(k + 1) + (labelName != null ? 1 : 0);
                createFreezePane(sxssfSheet, pane);
            }
            //  自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
            if (regionMap != null) {
                setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) regionMap.get(k + 1));
            }
            //  自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
            if (dropDownMap != null) {
                setDataValidation(sxssfSheet, (List<String[]>) dropDownMap.get(k + 1), listRow.size());
            }
            //  自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
            if (columnMap != null) {
                setColumnWidth(sxssfSheet, (HashMap) columnMap.get(k + 1));
            }
            //  默认样式。
            setStyle(cellStyle, font);

            CellStyle cell_style = null;
            CellStyle row_style = null;
            CellStyle column_style = null;
            //  写入小标题与数据。
            Integer SIZE = listRow.size() < MAX_ROWSUM ? listRow.size() : MAX_ROWSUM;
            Integer MAXSYTLE = listRow.size() < MAX_ROWSYTLE ? listRow.size() : MAX_ROWSYTLE;
            for (int i = 0; i < SIZE; i++) {
                sxssfRow = sxssfSheet.createRow(jRow);
                for (int j = 0; j < listRow.get(i).length; j++) {
                    Cell cell = createCell(sxssfRow, j, listRow.get(i)[j]);
                    cell.setCellStyle(cellStyle);
                    try {
                        //  自定义：每个表格每一列的样式（看该方法说明）。
                        //  样式过多会导致GC内存溢出！
                        if (columnStyles != null && jRow >= pane && i <= MAXSYTLE) {
                            if (jRow == pane && j == 0) {
                                column_style = cell.getRow().getSheet().getWorkbook().createCellStyle();
                            }
                            setExcelRowStyles(cell, column_style, wb, sxssfRow, (List) columnStyles.get(k + 1), j);
                        }
                        //  自定义：每个表格每一行的样式（看该方法说明）。
                        if (rowStyles != null && i <= MAXSYTLE) {
                            if (i == 0 && j == 0) {
                                row_style = cell.getRow().getSheet().getWorkbook().createCellStyle();
                            }
                            setExcelRowStyles(cell, row_style, wb, sxssfRow, (List) rowStyles.get(k + 1), jRow);
                        }
                        //  自定义：每一个单元格样式（看该方法说明）。
                        if (styles != null && i <= MAXSYTLE) {
                            if (i == 0) {
                                cell_style = cell.getRow().getSheet().getWorkbook().createCellStyle();
                            }
                            setExcelStyles(cell, cell_style, wb, sxssfRow, (List<List<Object[]>>) styles.get(k + 1), j, i);
                        }
                    } catch (Exception e) {
                        log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：The maximum number of cell styles was exceeded. You can define up to 4000 styles!");
                    }
                }
                jRow++;
            }
            k++;
        }
    }


    /**
     * @param cell         Cell对象。
     * @param wb           SXSSFWorkbook对象。
     * @param fontSize     字体大小。
     * @param bold         是否加粗。
     * @param center       是否左右上下居中。
     * @param isBorder     是否加边框
     * @param leftBoolean  左对齐
     * @param rightBoolean 右对齐
     * @param height       行高
     */
    public static void setExcelStyles(Cell cell, CellStyle cellStyle, SXSSFWorkbook wb, SXSSFRow sxssfRow, Integer fontSize, Boolean bold, Boolean center, Boolean isBorder, Boolean leftBoolean,
                                      Boolean rightBoolean, Integer fontColor, Integer height) {
        //保证了既可以新建一个CellStyle，又可以不丢失原来的CellStyle 的样式
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        //左右居中、上下居中
        if (center != null && center) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        //右对齐
        if (rightBoolean != null && rightBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        }
        //左对齐
        if (leftBoolean != null && leftBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
        }
        //边框
        if (isBorder != null && isBorder) {
            setBorder(cellStyle, isBorder);
        }
        //设置单元格字体样式
        XSSFFont font = (XSSFFont) wb.createFont();
        if (bold != null && bold) {
            font.setBold(bold);
        }
        //行高
        if (height != null) {
            sxssfRow.setHeight((short) (height * 2));
        }
        font.setFontName("宋体");
        font.setFontHeight(fontSize == null ? 12 : fontSize);
        cellStyle.setFont(font);
        //   点击可查看颜色对应的值： BLACK(8), WHITE(9), RED(10),
        font.setColor(IndexedColors.fromInt(fontColor == null ? 8 : fontColor).index);
        cell.setCellStyle(cellStyle);
    }


    public static void setExcelRowStyles(Cell cell, CellStyle cellStyle, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<Object[]> rowstyleList, int rowIndex) {
        if (rowstyleList != null && rowstyleList.size() > 0) {
            Integer[] rowstyle = (Integer[]) rowstyleList.get(1);
            for (int i = 0; i < rowstyle.length; i++) {
                if (rowIndex == rowstyle[i]) {
                    Boolean[] bool = (Boolean[]) rowstyleList.get(0);
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (rowstyleList.size() >= 3) {
                        int leng = rowstyleList.get(2).length;
                        if (leng >= 1) {
                            fontColor = (Integer) rowstyleList.get(2)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) rowstyleList.get(2)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) rowstyleList.get(2)[2];
                        }
                    }
                    setExcelStyles(cell, cellStyle, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                }
            }
        }
    }

    /**
     * 设置数据：无样式（行、列、单元格样式）
     *
     * @param wb
     * @param sxssfRow
     * @param dataLists
     * @param notBorderMap
     * @param regionMap
     * @param columnMap
     * @param paneMap
     * @param sheetName
     * @param labelName
     * @param dropDownMap
     * @throws Exception
     */
    public static void setDataListNoStyle(SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<String[]>> dataLists, HashMap notBorderMap, HashMap regionMap,
                                          HashMap columnMap, HashMap paneMap, String[] sheetName, String[] labelName, HashMap dropDownMap) throws Exception {
        if (dataLists == null) {
            log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：Export data(type:List<List<String[]>>) cannot be empty!");
        }
        if (sheetName == null) {
            log.debug("=== ===  === :Andyczy ExcelUtils Exception Message：Export sheet(type:String[]) name cannot be empty!");
        }
        int k = 0;
        for (List<String[]> listRow : dataLists) {
            SXSSFSheet sxssfSheet = wb.createSheet();
            wb.setSheetName(k, sheetName[k]);
            CellStyle cellStyle = wb.createCellStyle();
            XSSFFont font = (XSSFFont) wb.createFont();

            int jRow = 0;
            if (labelName != null) {
                //  自定义：大标题和样式。参数说明：new String[]{"表格数据一", "表格数据二", "表格数据三"}
                sxssfRow = sxssfSheet.createRow(0);
                Cell cell = createCell(sxssfRow, 0, labelName[k]);
                setMergedRegion(sxssfSheet, 0, 0, 0, listRow.get(0).length - 1);
                setLabelStyles(wb, cell, sxssfRow);
                jRow = 1;
            }
            //  自定义：每个表格固定表头（看该方法说明）。
            Integer pane = 1;
            if (paneMap != null && paneMap.get(k + 1) != null) {
                pane = (Integer) paneMap.get(k + 1) + (labelName != null ? 1 : 0);
                createFreezePane(sxssfSheet, pane);
            }
            //  自定义：每个单元格自定义合并单元格：对每个单元格自定义合并单元格（看该方法说明）。
            if (regionMap != null) {
                setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) regionMap.get(k + 1));
            }
            //  自定义：每个单元格自定义下拉列表：对每个单元格自定义下拉列表（看该方法说明）。
            if (dropDownMap != null) {
                setDataValidation(sxssfSheet, (List<String[]>) dropDownMap.get(k + 1), listRow.size());
            }
            //  自定义：每个表格自定义列宽：对每个单元格自定义列宽（看该方法说明）。
            if (columnMap != null) {
                setColumnWidth(sxssfSheet, (HashMap) columnMap.get(k + 1));
            }
            //  默认样式。
            setStyle(cellStyle, font);

            //  写入小标题与数据。
            Integer SIZE = listRow.size() < MAX_ROWSUM ? listRow.size() : MAX_ROWSUM;
            for (int i = 0; i < SIZE; i++) {
                sxssfRow = sxssfSheet.createRow(jRow);
                for (int j = 0; j < listRow.get(i).length; j++) {
                    Cell cell = createCell(sxssfRow, j, listRow.get(i)[j]);
                    cell.setCellStyle(cellStyle);
                }
                jRow++;
            }
            k++;
        }
    }



    public static void writeAndColse(SXSSFWorkbook sxssfWorkbook, OutputStream outputStream) throws Exception {
        try {
            if (outputStream != null) {
                sxssfWorkbook.write(outputStream);
                sxssfWorkbook.dispose();
                outputStream.flush();
                outputStream.close();
            }
        } catch (Exception e) {
            System.out.println(" Andyczy ExcelUtils Exception Message：Output stream is not empty !");
            e.getSuppressed();
        }
    }

    /**
     * 功能描述：所有自定义单元格样式
     * 使用的方法：是否居中？，是否右对齐？，是否左对齐？， 是否加粗？，是否有边框？  —— 颜色、字体、行高？
     * HashMap cellStyles = new HashMap();
     * List< List<Object[]>> list = new ArrayList<>();
     * List<Object[]> objectsList = new ArrayList<>();
     * List<Object[]> objectsListTwo = new ArrayList<>();
     * objectsList.add(new Boolean[]{true, false, false, false, true});      //1、样式一（必须放第一）
     * objectsList.add(new Integer[]{10, 12});                               //1、颜色值 、字体大小、行高（必须放第二）
     * <p>
     * objectsListTwo.add(new Boolean[]{false, false, false, true, true});   //2、样式二（必须放第一）
     * objectsListTwo.add(new Integer[]{10, 12,null});                       //2、颜色值 、字体大小、行高（必须放第二）
     * <p>
     * objectsList.add(new Integer[]{5, 1});                                 //1、第五行第一列
     * objectsList.add(new Integer[]{6, 1});                                 //1、第六行第一列
     * <p>
     * objectsListTwo.add(new Integer[]{2, 1});                              //2、第二行第一列
     * <p>
     * cellStyles.put(1, list);                                              //第一个表格所有自定义单元格样式
     *
     * @param cell
     * @param wb
     * @param styles
     */
    public static void setExcelStyles(Cell cell, CellStyle cellStyle, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<Object[]>> styles, int cellIndex, int rowIndex) {
        if (styles != null) {
            for (int z = 0; z < styles.size(); z++) {
                List<Object[]> stylesList = styles.get(z);
                if (stylesList != null) {
                    //样式
                    Boolean[] bool = (Boolean[]) stylesList.get(0);
                    //颜色和字体
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    //当有设置颜色值 、字体大小、行高才获取值
                    if (stylesList.size() >= 2) {
                        int leng = stylesList.get(1).length;
                        if (leng >= 1) {
                            fontColor = (Integer) stylesList.get(1)[0];
                        }
                        if (leng >= 2) {
                            fontSize = (Integer) stylesList.get(1)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) stylesList.get(1)[2];
                        }
                    }
                    //第几行第几列
                    for (int m = 2; m < stylesList.size(); m++) {
                        Integer[] str = (Integer[]) stylesList.get(m);
                        if (cellIndex + 1 == (str[1]) && rowIndex + 1 == (str[0])) {
                            setExcelStyles(cell, cellStyle, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                        }
                    }
                }
            }
        }
    }


    /**
     * 大标题样式
     *
     * @param wb
     * @param cell
     * @param sxssfRow
     */
    public static void setLabelStyles(SXSSFWorkbook wb, Cell cell, SXSSFRow sxssfRow) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorder(cellStyle, true);
        sxssfRow.setHeight((short) (399 * 2));
        XSSFFont font = (XSSFFont) wb.createFont();
        font.setFontName("宋体");
        font.setFontHeight(16);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 默认样式
     *
     * @param cellStyle
     * @param font
     * @return
     */
    public static void setStyle(CellStyle cellStyle, XSSFFont font) {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        font.setFontName("宋体");
        cellStyle.setFont(font);
        font.setFontHeight(12);
        setBorder(cellStyle, true);
    }


    /**
     * 判断字符串是否为空
     *
     * @param str
     * @return
     */
    public static boolean isBlank(String str) {
        int strLen;
        if (str != null && (strLen = str.length()) != 0) {
            for (int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(str.charAt(i))) {
                    return false;
                }
            }
            return true;
        } else {
            return true;
        }
    }

    /**
     * 功能描述: 锁定行（固定表头）
     * 参数说明：
     * HashMap setPaneMap = new HashMap();
     * //第一个表格、第三行开始固定表头
     * setPaneMap.put(1,3);
     *
     * @param sxssfSheet
     * @param row
     */
    public static void createFreezePane(SXSSFSheet sxssfSheet, Integer row) {
        if (row != null && row > 0) {
            sxssfSheet.createFreezePane(0, row, 0, 1);
        }
    }

    /**
     * 功能描述: 自定义列宽
     * 参数说明：
     * HashMap<Integer, HashMap<Integer, Integer>> columnMap = new HashMap<>();
     * HashMap<Integer, Integer> mapColumn = new HashMap<>();
     * //第一列、宽度为 3[3的大小就是两个12号字体刚刚好的列宽]（注意：excel从零行开始数）
     * mapColumn.put(0, 3);
     * mapColumn.put(1, 20);
     * mapColumn.put(2, 15);
     * //第一个单元格列宽
     * columnMap.put(1, mapColumn);
     *
     * @param sxssfSheet
     * @param map
     */
    public static void setColumnWidth(SXSSFSheet sxssfSheet, HashMap map) {
        if (map != null) {
            Iterator iterator = map.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry entry = (Map.Entry) iterator.next();
                Object key = entry.getKey();
                Object val = entry.getValue();
                sxssfSheet.setColumnWidth((int) key, (int) val * 512);
            }
        }
    }


    /**
     * 功能描述: excel 合并单元格
     * 参数说明：
     * List<List<Integer[]>> regionMap = new ArrayList<>();
     * List<Integer[]> regionList = new ArrayList<>();
     * //代表起始行号，终止行号， 起始列号，终止列号进行合并。（注意：excel从零行开始数）
     * regionList.add(new Integer[]{1, 1, 0, 10});
     * regionList.add(new Integer[]{2, 3, 1, 1});
     * //第一个表格设置。
     * regionMap.put(1, regionList);
     *
     * @param sheet
     * @param rowColList
     */
    public static void setMergedRegion(SXSSFSheet sheet, ArrayList<Integer[]> rowColList) {
        if (rowColList != null && rowColList.size() > 0) {
            for (int i = 0; i < rowColList.size(); i++) {
                Integer[] str = rowColList.get(i);
                if (str.length > 0 && str.length == 4) {
                    Integer firstRow = str[0];
                    Integer lastRow = str[1];
                    Integer firstCol = str[2];
                    Integer lastCol = str[3];
                    setMergedRegion(sheet, firstRow, lastRow, firstCol, lastCol);
                }
            }
        }
    }

    /**
     * 功能描述: 合并单元格
     *
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public static void setMergedRegion(SXSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }


    /**
     * 功能描述:下拉列表
     * 参数说明：
     * HashMap dropDownMap = new HashMap();
     * List<String[]> dropList = new ArrayList<>();
     * //必须放第一：设置下拉列表的列（excel从零行开始数）
     * String[] sheetDropData = new String[]{"1", "2", "4"};
     * //下拉的值放在 sheetDropData 后面。
     * String[] sex = {"男,女"};
     * dropList.add(sheetDropData);
     * dropList.add(sex);
     * //第一个表格设置。
     * dropDownMap.put(1,dropList);
     *
     * @param sheet
     * @param dropDownListData
     * @param dataListSize
     */
    public static void setDataValidation(SXSSFSheet sheet, List<String[]> dropDownListData, int dataListSize) {
        if (dropDownListData.size() > 0) {
            for (int col = 0; col < dropDownListData.get(0).length; col++) {
                Integer colv = Integer.parseInt(dropDownListData.get(0)[col]);
                setDataValidation(sheet, dropDownListData.get(col + 1), 1, dataListSize < 100 ? 500 : dataListSize, colv, colv);
            }
        }
    }


    /**
     * 功能描述:下拉列表
     *
     * @param xssfWsheet
     * @param list
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public static void setDataValidation(SXSSFSheet xssfWsheet, String[] list, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
        DataValidationHelper helper = xssfWsheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(list);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.createErrorBox(DataValidationError1, DataValidationError2);
        //  处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        xssfWsheet.addValidationData(dataValidation);
    }

    /**
     * 功能描述：设置边框
     *
     * @param cellStyle
     * @param isBorder
     */
    public static void setBorder(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
        }
    }

    /**
     * 验证是否是日期
     *
     * @param strDate
     * @return
     */
    public static boolean verificationDate(String strDate, String style) {
        Date date = null;
        if (style == null) {
            style = "yyyy-MM-dd";
        }
        SimpleDateFormat formatter = new SimpleDateFormat(style);
        try {
            formatter.parse(strDate);
        } catch (Exception e) {
            return false;
        }
        return true;
    }

    public static String strToDateFormat(String strDate, String style, String expectDateFormatStr) {
        Date date = null;
        if (style == null) {
            style = "yyyy-MM-dd";
        }
        // 日期字符串转成date类型
        SimpleDateFormat formatter = new SimpleDateFormat(style);
        try {
            date = formatter.parse(strDate);
        } catch (Exception e) {
            return null;
        }
        // 转成指定的日期格式
        SimpleDateFormat sdf = new SimpleDateFormat(expectDateFormatStr == null ? style : expectDateFormatStr);
        String str = sdf.format(date);
        return str;
    }


}
