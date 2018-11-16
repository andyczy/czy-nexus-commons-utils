package com.github.andyczy.java.excel;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFont;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.util.CellUtil.createCell;


public class ExcelUtils {


    private static final ThreadLocal<SimpleDateFormat> fmt = new ThreadLocal<>();
    private static final String MESSAGE_FORMAT = "yyyy-MM-dd";

    private static final ThreadLocal<DecimalFormat> df = new ThreadLocal<>();
    private static final String MESSAGE_FORMAT_df = "#.###";
    private static final String DataValidationError1 = "This system - remind you!";
    private static final String DataValidationError2 = "The data is not standardized, please select the data in the table list!";

    private static final SimpleDateFormat getDateFormat() {
        SimpleDateFormat format = fmt.get();
        if (format == null) {
            format = new SimpleDateFormat(MESSAGE_FORMAT, Locale.getDefault());
            fmt.set(format);
        }
        return format;
    }

    private static final DecimalFormat getDecimalFormat() {
        DecimalFormat format = df.get();
        if (format == null) {
            format = new DecimalFormat(MESSAGE_FORMAT_df);
            df.set(format);
        }
        return format;
    }



    public static Boolean exportForExcel(HttpServletResponse response, List<List<String[]>> dataLists, HashMap notBorderMap,
                                         HashMap regionMap, HashMap columnMap, HashMap styles, HashMap paneMap, String fileName,
                                         String[] sheetName, String[] labelName, HashMap rowStyles, HashMap columnStyles, HashMap dropDownMap) {
        long startTime = System.currentTimeMillis();
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(1000);
        OutputStream outputStream = null;
        SXSSFRow sxssfRow = null;
        try {
            int k = 0;
            for (List<String[]> list : dataLists) {
                SXSSFSheet sxssfSheet = sxssfWorkbook.createSheet();
                sxssfSheet.setDefaultColumnWidth((short) 16);
                sxssfWorkbook.setSheetName(k, sheetName[k]);

                int jRow = 0;
                if (labelName != null) {
                    sxssfRow = sxssfSheet.createRow(jRow);
                    Cell cell = createCell(sxssfRow, jRow, labelName[k]);
                    setMergedRegion(sxssfSheet, 0, 0, 0, list.get(0).length - 1);
                    setExcelStyles(cell, sxssfWorkbook, sxssfRow, 16, true, true, false, false, false, null, 399);
                    jRow = 1;
                }
                if (regionMap != null) {
                    setMergedRegion(sxssfSheet, (ArrayList<Integer[]>) regionMap.get(k + 1));
                }
                if (dropDownMap != null) {
                    setDataValidation(sxssfSheet, (List<String[]>) dropDownMap.get(k + 1), list.size());
                }
                if (columnMap != null) {
                    setColumnWidth(sxssfSheet, (HashMap) columnMap.get(k + 1));
                }
                Integer pane = 1;
                if (paneMap != null) {
                    pane = (Integer) paneMap.get(k + 1) + (labelName != null ? 1 : 0);
                    createFreezePane(sxssfSheet, pane);
                }

                for (String[] listValue : list) {
                    int columnIndex = 0;
                    sxssfRow = sxssfSheet.createRow(jRow);
                    for (int j = 0; j < listValue.length; j++) {
                        Cell cells = sxssfRow.createCell(j);
                        Cell cell = createCell(sxssfRow, columnIndex, listValue[j]);
                        columnIndex++;
                        setExcelStyles(notBorderMap, cell, sxssfWorkbook, sxssfRow, k, jRow);
                        if (columnStyles != null && jRow >= pane) {
                            setExcelCellStyles(cell, sxssfWorkbook, sxssfRow, (List) columnStyles.get(k + 1), j);
                        }
                        if (rowStyles != null) {
                            setExcelCellStyles(cell, sxssfWorkbook, sxssfRow, (List) rowStyles.get(k + 1), jRow);
                        }
                        if (styles != null) {
                            setExcelStyles(cells, sxssfWorkbook, sxssfRow, (List<List<Object[]>>) styles.get(k + 1), j, jRow);
                        }
                    }
                    jRow++;
                }
                k++;
            }
            response.setHeader("Charset", "UTF-8");
            response.setHeader("Content-Type", "application/force-download");
            response.setHeader("Content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(fileName == null ? "Excel" : sheetName[0], "utf8") + ".xlsx");
            response.flushBuffer();
            outputStream = response.getOutputStream();
            sxssfWorkbook.write(outputStream);
            sxssfWorkbook.dispose();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("======= Excel Run Time:  " + (System.currentTimeMillis() - startTime) + " ms");
        return true;
    }


    @SuppressWarnings({"deprecation", "rawtypes"})
    public static List<List<LinkedHashMap<String, String>>> importForExcelData(Workbook book, String[] sheetName, HashMap indexMap, HashMap continueRowMap) {
        long startTime = System.currentTimeMillis();
        try {
            List<List<LinkedHashMap<String, String>>> returnDataList = new ArrayList<>();
            for (int k = 0; k <= sheetName.length - 1; k++) {
                Sheet sheet = book.getSheetAt(k);
                int rowCount = sheet.getLastRowNum() + 1;
                Row valueRow = null;

                List<LinkedHashMap<String, String>> rowListValue = new ArrayList<>();
                LinkedHashMap<String, String> cellHashMap = null;

                int irow = 1;
                if (indexMap != null && indexMap.get(k + 1) != null) {
                    irow = Integer.valueOf(indexMap.get(k + 1).toString()) - 1;
                }
                for (int i = irow <= 0 ? 1 : irow; i < rowCount; i++) {
                    valueRow = sheet.getRow(i);
                    if (valueRow == null) {
                        continue;
                    }
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

                    for (int j = 0; j < valueRow.getLastCellNum(); j++) {
                        cellHashMap.put(Integer.toString(j), getCellVal(valueRow.getCell(j)));
                    }
                    if (cellHashMap.size() > 0) {
                        rowListValue.add(cellHashMap);
                    }
                }
                returnDataList.add(rowListValue);
            }
            System.out.println("=======  Excel Run Time:  " + (System.currentTimeMillis() - startTime) + " ms");
            return returnDataList;
        } catch (Exception e) {
            System.out.println("=======  Excel Exception");
            e.printStackTrace();
            return null;
        }
    }





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


    public static void createFreezePane(SXSSFSheet sxssfSheet, Integer row) {
        if (row != null && row > 0) {
            sxssfSheet.createFreezePane(0, row, 0, 1);
        }
    }


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

    public static void setMergedRegion(SXSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }



    public static void setDataValidation(SXSSFSheet sheet, List<String[]> dropDownListData, int dataListSize) {
        if (dropDownListData.size() > 0) {
            for (int col = 0; col < dropDownListData.get(0).length; col++) {
                Integer colv = Integer.parseInt(dropDownListData.get(0)[col]);
                setDataValidation(sheet, dropDownListData.get(col + 1), 1, dataListSize < 100 ? 500 : dataListSize, colv, colv);
            }
        }
    }


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


    public static void setExcelStyles(HashMap notBorderMap, Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, int k, int jRow) {
        Boolean border = true;
        if (notBorderMap != null) {
            Integer[] borderInt = (Integer[]) notBorderMap.get(k + 1);
            for (int i = 0; i < borderInt.length; i++) {
                if (borderInt[i] == jRow) {
                    border = false;
                }
            }
        }
        setExcelStyles(cell, wb, sxssfRow, null, null, true, border, false, false, null, null);
    }

    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, Integer fontSize, Boolean bold, Boolean center, Boolean isBorder, Boolean leftBoolean,
                                      Boolean rightBoolean, Integer fontColor, Integer height) {
        CellStyle cellStyle = wb.createCellStyle();
        if (center != null && center) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        if (rightBoolean != null && rightBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        }
        if (leftBoolean != null && leftBoolean) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
        }
        if (isBorder != null && isBorder) {
            setBorder(cellStyle, isBorder);
        }
        XSSFFont font = (XSSFFont) wb.createFont();
        if (bold != null && bold) {
            font.setBold(bold);
        }
        if (height != null) {
            sxssfRow.setHeight((short) (height * 2));
        }
        font.setFontName("宋体");
        font.setFontHeight(fontSize == null ? 12 : fontSize);
        cellStyle.setFont(font);
        font.setColor(IndexedColors.fromInt(fontColor == null ? 8 : fontColor).index);
        cell.setCellStyle(cellStyle);
    }

    public static void setExcelCellStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<Object[]> rowstyleList, int rowIndex) {
        if (rowstyleList != null && rowstyleList.size() > 0) {
            Integer[] rowstyle = (Integer[]) rowstyleList.get(1);
            for (int i = 0; i < rowstyle.length; i++) {
                if (rowIndex == rowstyle[i]) {
                    Boolean[] bool = (Boolean[]) rowstyleList.get(0);
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    if (rowstyleList.size() >= 3) {
                        int leng = rowstyleList.get(2).length;
                        fontColor = (Integer) rowstyleList.get(2)[0];
                        if (leng >= 2) {
                            fontSize = (Integer) rowstyleList.get(2)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) rowstyleList.get(2)[2];
                        }
                    }
                    setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                }
            }
        }
    }

    public static void setExcelStyles(Cell cell, SXSSFWorkbook wb, SXSSFRow sxssfRow, List<List<Object[]>> styles, int cellIndex, int rowIndex) {
        if (styles != null) {
            for (int z = 0; z < styles.size(); z++) {
                List<Object[]> stylesList = styles.get(z);
                if (stylesList != null) {
                    Boolean[] bool = (Boolean[]) stylesList.get(0);
                    Integer fontColor = null;
                    Integer fontSize = null;
                    Integer height = null;
                    if (stylesList.size() >= 2) {
                        int leng = stylesList.get(1).length;
                        fontColor = (Integer) stylesList.get(1)[0];
                        if (leng >= 2) {
                            fontSize = (Integer) stylesList.get(1)[1];
                        }
                        if (leng >= 3) {
                            height = (Integer) stylesList.get(1)[2];
                        }
                    }
                    for (int m = 2; m < stylesList.size(); m++) {
                        Integer[] str = (Integer[]) stylesList.get(m);
                        if (cellIndex + 1 == (str[1]) && rowIndex + 1 == (str[0])) {
                            setExcelStyles(cell, wb, sxssfRow, fontSize, Boolean.valueOf(bool[3]), Boolean.valueOf(bool[0]), Boolean.valueOf(bool[4]), Boolean.valueOf(bool[2]), Boolean.valueOf(bool[1]), fontColor, height);
                        }
                    }
                }
            }
        }
    }


    private static void setBorder(CellStyle cellStyle, Boolean isBorder) {
        if (isBorder) {
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
        }
    }



    public static String getCellVal(Cell cell) {
        String val = null;
        if (cell != null) {
            CellType cellType = cell.getCellTypeEnum();
            switch (cellType) {
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        val = getDateFormat().format(cell.getDateCellValue());
                    } else {
                        val = getDecimalFormat().format(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    val = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    val = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK:
                    val = cell.getStringCellValue();
                    break;
                case ERROR:
                    val = "ERROR!";
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

}

