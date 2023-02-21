package com.fastchar.template.provider;

import com.fastchar.core.FastHandler;
import com.fastchar.template.FastTemplateHelper;
import com.fastchar.template.info.FastWordTableInfo;
import com.fastchar.template.interfaces.IFastTemplateRender;
import com.fastchar.utils.FastFileUtils;
import com.fastchar.utils.FastNumberUtils;
import com.fastchar.utils.FastStringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word模板渲染器，变量格式${*} 注意：为了避免word分割格式${*} 请使用复制粘贴的方式到word中！！
 *
 * @author 沈建（Janesen）
 * @date 2021/12/6 17:11
 */
public class FastWordTemplateRender implements IFastTemplateRender {
    private static final Pattern PLACE_HOLDER_PATTERN = Pattern.compile("(\\$\\{(.*)})", Pattern.DOTALL);
    private static final Pattern PLACE_HOLDER_LIST_PATTERN = Pattern.compile("([^${}]*)\\[i]",Pattern.DOTALL);

    @Override
    public void onRender(FastHandler handler, InputStream templateInputStream, OutputStream newFileOutStream) {
        if (!FastFileUtils.isWordFile(handler.getString("__fileName"))) {
            return;
        }
        handler.setCode(0);
        XWPFDocument document = null;
        try {
            document = new XWPFDocument(OPCPackage.open(templateInputStream));
            wrapListTable(handler, document);
            renderNormal(handler, document);
            renderTables(handler, document);
            document.write(newFileOutStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (document != null) {
                    document.close();
                }
                if (newFileOutStream != null) {
                    newFileOutStream.close();
                }
            } catch (Exception ignored) {
            }
        }
    }

    private void wrapListTable(FastHandler handler, XWPFDocument document) {
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            Map<Integer, String> cellPatternMap = new HashMap<>();
            Map<Integer, XWPFTableCell> cellMap = new HashMap<>();

            int maxRowData = 0;
            int maxCell = 0;

            for (XWPFTableRow row : table.getRows()) {
                List<XWPFTableCell> tableCells = row.getTableCells();

                for (int i = 0; i < tableCells.size(); i++) {
                    maxCell = Math.max(maxCell, i);

                    XWPFTableCell cell = tableCells.get(i);
                    String text = cell.getText();
                    Matcher matcher = PLACE_HOLDER_PATTERN.matcher(text);

                    boolean hasList = false;
                    while (matcher.find()) {
                        String key = matcher.group(2).trim();
                        if (key.contains("[i]")) {
                            Matcher listMatcher = PLACE_HOLDER_LIST_PATTERN.matcher(key);
                            while (listMatcher.find()) {
                                maxRowData = Math.max(maxRowData, FastNumberUtils.formatToInt(FastTemplateHelper.renderData(handler, listMatcher.group(1) + ".length")));
                            }
                            hasList = true;
                        }
                    }
                    if (hasList) {
                        cell.setText(text.replace("[i]", "[0]"));
                        cellPatternMap.put(i, text);
                        cellMap.put(i, cell);
                    }
                }
            }

            for (int i = 1; i < maxRowData; i++) {
                XWPFTableRow row = table.createRow();
                for (int i1 = 0; i1 <= maxCell; i1++) {
                    XWPFTableCell cell = row.getCell(i1);
                    if (cell == null) {
                        cell = row.createCell();
                    }
                    String cellPattern = cellPatternMap.get(i1);
                    if (FastStringUtils.isEmpty(cellPattern)) {
                        cellPattern = "";
                    }

                    XWPFTableCell sourceCell = cellMap.get(i1);
                    if (sourceCell != null) {
                        cell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
                    }

                    cell.setText(cellPattern.replace("[i]", "[" + i + "]"));
                }
            }
        }
    }

    private void renderNormal(FastHandler handler, XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            replaceParagraph(handler, paragraph);
        }
    }

    private void renderTables(FastHandler handler, XWPFDocument document) {
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            for (XWPFTableRow row : table.getRows()) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells) {
                    for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                        replaceParagraph(handler, paragraph);
                    }
                }
            }
        }
    }


    private void replaceParagraph(FastHandler handler, XWPFParagraph paragraph) {
        List<XWPFRun> unMatchedRun = new ArrayList<>();
        for (XWPFRun run : paragraph.getRuns()) {
            int textPosition = Math.max(run.getTextPosition(), 0);
            String text = toText(unMatchedRun) + run.getText(textPosition);
            if (FastStringUtils.isEmpty(text)) {
                continue;
            }
            boolean hasChange = false;
            Matcher matcher = PLACE_HOLDER_PATTERN.matcher(text);
            while (matcher.find()) {
                hasChange = true;
                String whole = matcher.group(1);
                String key = matcher.group(2).trim();

                Object invokeValue = FastTemplateHelper.renderData(handler, key);
                if (invokeValue == null) {
                    invokeValue = "";
                }

                if (invokeValue instanceof FastWordTableInfo) {
                    createTable(paragraph, (FastWordTableInfo) invokeValue);
                    text = "";
                    continue;
                }
                text = text.replace(whole, invokeValue.toString());
            }
            if (hasChange) {
                clearText(unMatchedRun);
                unMatchedRun.clear();
                run.setText(text, textPosition);
                formatBreakLine(run);
            }
            unMatchedRun.add(run);
        }
    }

    private String toText(List<XWPFRun> runs) {
        StringBuilder stringBuilder = new StringBuilder();
        for (XWPFRun run : runs) {
            int textPosition = Math.max(run.getTextPosition(), 0);
            stringBuilder.append(run.getText(textPosition));
        }
        return stringBuilder.toString();
    }

    private void clearText(List<XWPFRun> runs) {
        for (XWPFRun run : runs) {
            int textPosition = Math.max(run.getTextPosition(), 0);
            run.setText("", textPosition);
        }
    }

    private void formatBreakLine(XWPFRun run) {
        if (run.getText(0) != null && run.getText(0).contains("\n")) {
            String[] lines = run.getText(0).split("\n");
            if (lines.length > 0) {
                run.setText(lines[0], 0);
                for (int i = 1; i < lines.length; i++) {
                    run.addBreak();
                    run.setText(lines[i]);
                }
            }
        }
    }


    private void createTable(XWPFParagraph paragraph, FastWordTableInfo wordTableInfo) {

        if (wordTableInfo.getValues() == null || wordTableInfo.getValues().isEmpty()) {
            return;
        }

        int rows = wordTableInfo.getValues().size();
        int cols = wordTableInfo.getValues().get(0).size();
        if (wordTableInfo.getTitles() != null && wordTableInfo.getTitles().size() > 0) {
            rows += 1;
            cols = wordTableInfo.getTitles().size();
        }

        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable xwpfTable = paragraph.getDocument().insertNewTbl(cursor);

        for (int i = 1; i < cols; i++) {
            xwpfTable.getRow(0).createCell();
        }

        for (int i = 1; i < rows; i++) {
            xwpfTable.createRow();
        }

        fillTable(xwpfTable, wordTableInfo);
    }

    private void fillTable(XWPFTable table, FastWordTableInfo tableInfo) {

        int beginRow = 0;
        if (tableInfo.getTitles() != null) {
            for (int i = 0; i < tableInfo.getTitles().size(); i++) {
                table.getRow(0).getCell(i).setText(tableInfo.getTitles().get(i));
                beginRow = 1;
            }
        }

        for (int i = 0; i < tableInfo.getValues().size(); i++) {
            XWPFTableRow row = table.getRow(beginRow + i);
            List<Object> values = tableInfo.getValues().get(i);
            for (int i1 = 0; i1 < values.size(); i1++) {
                row.getCell(i1).setText(String.valueOf(values.get(i1)));
            }
        }
    }

}
