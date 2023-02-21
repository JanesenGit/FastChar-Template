package com.fastchar.template.provider;

import com.fastchar.core.FastChar;
import com.fastchar.core.FastHandler;
import com.fastchar.template.FastTemplateHelper;
import com.fastchar.template.interfaces.IFastTemplateRender;
import com.fastchar.utils.FastDateUtils;
import com.fastchar.utils.FastFileUtils;
import com.fastchar.utils.FastNumberUtils;
import com.fastchar.utils.FastStringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel文件模板渲染器
 *
 * @author 沈建（Janesen）
 * @date 2021/12/6 17:37
 */
public class FastExcelTemplateRender implements IFastTemplateRender {
    @Override
    public void onRender(FastHandler handler, InputStream templateInputStream, OutputStream newFileOutStream) {
        if (!FastFileUtils.isExcelFile(handler.getString("__fileName"))) {
            return;
        }
        handler.setCode(0);
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(templateInputStream);
            wrapList(handler, workbook);
            renderNormal(handler, workbook);
            workbook.write(newFileOutStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (newFileOutStream != null) {
                    newFileOutStream.close();
                }
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception ignored) {
            }
        }
    }


    private void wrapList(FastHandler handler, Workbook workbook) {
        String regStr = "\\$\\{(.*)}";
        Pattern compile = Pattern.compile(regStr);


        int sheetCount = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int rowNum = 0; rowNum < rowCount; rowNum++) {
                Row dataRow = sheet.getRow(rowNum);
                if (dataRow == null) {
                    continue;
                }

                int maxRowData = 0;
                int maxCell = 0;
                CellStyle parentStyle = null;

                Map<Integer, String> cellPatternMap = new HashMap<>();

                int cellCount = dataRow.getLastCellNum();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    maxCell = Math.max(cellNum, maxCell);

                    Cell cell = dataRow.getCell(cellNum);
                    Object value = getCellValue(workbook, cell);
                    if (value != null) {
                        String text = value.toString();
                        Matcher matcher = compile.matcher(text);

                        boolean hasList = false;
                        while (matcher.find()) {
                            if (matcher.group(1).contains("[i]")) {
                                String listRegStr = "([^${}]*)\\[i]";
                                Pattern listCompile = Pattern.compile(listRegStr);
                                Matcher listMatcher = listCompile.matcher(matcher.group(1));
                                while (listMatcher.find()) {
                                    maxRowData = Math.max(maxRowData, FastNumberUtils.formatToInt(FastTemplateHelper.renderData(handler, listMatcher.group(1) + ".length")));
                                }
                                parentStyle = cell.getCellStyle();
                                hasList = true;
                            }
                        }
                        if (hasList) {
                            cell.setCellValue(text.replace("[i]", "[0]"));
                            cellPatternMap.put(cellNum, text);
                        }
                    }
                }

                int lastInsertRow = rowNum + 1;
                for (int newRowNum = 1; newRowNum < maxRowData; newRowNum++) {
                    sheet.shiftRows(lastInsertRow, lastInsertRow + 1, 1);
                    Row row = sheet.getRow(lastInsertRow);
                    if (row == null) {
                        row = sheet.createRow(lastInsertRow);
                    }

                    for (int cellNum = 0; cellNum <= maxCell; cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        if (cell == null) {
                            cell = row.createCell(cellNum);
                        }
                        if (parentStyle != null) {
                            cell.setCellStyle(parentStyle);
                        }

                        String cellPattern = cellPatternMap.get(cellNum);
                        if (FastStringUtils.isEmpty(cellPattern)) {
                            cellPattern = "";
                        }
                        cell.setCellValue(cellPattern.replace("[i]", "[" + newRowNum + "]"));
                    }

                    lastInsertRow = lastInsertRow + 1;
                }
                if (maxRowData > 0) {
                    rowNum = 0;
                    rowCount = sheet.getPhysicalNumberOfRows();
                }
            }
        }
    }

    private void renderNormal(FastHandler handler, Workbook workbook) {
        int sheetCount = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int i1 = 0; i1 < rowCount; i1++) {
                Row dataRow = sheet.getRow(i1);
                if (dataRow == null) {
                    continue;
                }
                int cellCount = dataRow.getLastCellNum();
                for (int i2 = 0; i2 < cellCount; i2++) {
                    Cell cell = dataRow.getCell(i2);
                    replaceCell(handler, workbook, cell);
                }
            }
        }
    }


    private Object getCellValue(Workbook workbook, Cell cell) {
        if (workbook == null) {
            return null;
        }
        if (cell == null) {
            return null;
        }
        return takeCellValue(workbook, cell, cell.getCellType());
    }

    private Object takeCellValue(Workbook workbook, Cell cell, CellType cellType) {
        try {
            if (workbook == null) {
                return null;
            }
            if (cell == null) {
                return null;
            }
            if (cellType == CellType.BLANK || cellType == CellType.STRING) {
                return cell.getStringCellValue();
            } else if (cellType == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    return FastDateUtils.format(cell.getDateCellValue(), FastChar.getConstant().getDateFormat());
                } else {
                    return NumberToTextConverter.toText(cell.getNumericCellValue());
                }
            } else if (cellType == CellType.BOOLEAN) {
                return cell.getBooleanCellValue();
            } else if (cellType == CellType.FORMULA) {
                return takeCellValue(workbook, cell, cell.getCachedFormulaResultType());
            }
            return cell.getStringCellValue();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private void replaceCell(FastHandler handler, Workbook workbook, Cell cell) {
        Object value = getCellValue(workbook, cell);
        if (value != null) {
            String regStr = "\\$\\{(.*)}";
            Pattern compile = Pattern.compile(regStr);
            String text = value.toString();
            Matcher matcher = compile.matcher(text);
            while (matcher.find()) {
                String key = matcher.group(1);
                Object invokeValue = FastTemplateHelper.renderData(handler, key);
                if (invokeValue == null) {
                    invokeValue = "";
                }
                text = text.replace("${" + key + "}", invokeValue.toString());
            }
            cell.setCellValue(text);
        }
    }
}
