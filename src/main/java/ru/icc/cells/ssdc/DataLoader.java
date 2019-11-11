/*
 * Copyright 2015-17 Alexey O. Shigarov (shigarov@gmail.com)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package ru.icc.cells.ssdc;

import java.io.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import ru.icc.cells.ssdc.model.*;
import ru.icc.cells.ssdc.model.style.*;

public final class DataLoader {
    private File sourceWorkbookFile;
    private Workbook workbook;
    private Sheet sheet;
    private String sheetName;
    private int rowIndex;
    private int cellCount;
    private CPoint cellShift = null;

    private static final String REF_POINT_VAL = "$START";
    private static final String END_POINT_VAL = "$END";
    //private static final String TBL_NAME      = "$NAME";
    //private static final String TBL_MEASURE   = "$MEASURE";

    private boolean withoutSuperscript;

    public boolean isWithoutSuperscript() {
        return withoutSuperscript;
    }

    public void setWithoutSuperscript(boolean withoutSuperscript) {
        this.withoutSuperscript = withoutSuperscript;
    }

    private boolean useCellValue;

    public boolean isUseCellValue() {
        return useCellValue;
    }

    public void setUseCellValue(boolean useCellValue) {
        this.useCellValue = useCellValue;
    }

    public void loadWorkbook(File excelFile) throws IOException {
        FileInputStream fin = new FileInputStream(excelFile);
        workbook = new XSSFWorkbook(fin);
        formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        sourceWorkbookFile = excelFile;
    }

    private void reset() {
        rowIndex = 0;
    }

    public int numOfSheets() {
        if (null == workbook)
            throw new IllegalStateException("The workbook is not loaded");
        return workbook.getNumberOfSheets();
    }

    public void goToSheet(int index) {
        if (null == workbook)
            throw new IllegalStateException("The workbook is not loaded");
        sheet = workbook.getSheetAt(index);
        sheetName = sheet.getSheetName();
        reset();
    }

    public String getCurrentSheetName() {
        return sheetName;
    }

    public CTable nextTable() {
        if (null == sheet)
            throw new IllegalStateException("The sheet is not initialized");

        CPoint refPnt = findRefPoint(sheet, rowIndex);
        if (null == refPnt) return null;
        cellShift = refPnt;

        CPoint endPnt = findEndPoint(sheet, refPnt.r);
        if (null == endPnt) return null;

        int numOfCols = endPnt.c - refPnt.c + 1;
        int numOfRows = endPnt.r - refPnt.r + 1;
        CTable table = new CTable(numOfRows, numOfCols);

        //CCell cell;
        Cell excelCell;
        Row row = null;
        CellRangeAddress cellRangeAddress = null;
        boolean isCell = false;

        int refRowAdr = refPnt.r;
        int endRowAdr = endPnt.r;
        int refColAdr = refPnt.c;
        int endColAdr = endPnt.c;

        for (int i = refRowAdr; i <= endRowAdr; i++) {
            row = sheet.getRow(i);

            // TODO надо внимательнее разобраться со случаем, когда r == null
            if (null == row)
                continue;

            for (int j = refColAdr; j <= endColAdr; j++) {
                // TODO надо внимательнее разобраться со случаем, когда excelCell == null
                excelCell = row.getCell(j, Row.CREATE_NULL_AS_BLANK);

                int colAdr = excelCell.getColumnIndex() - refColAdr + 1;
                int rowAdr = excelCell.getRowIndex() - refRowAdr + 1;

                int cl = colAdr;
                int cr = colAdr;
                int rt = rowAdr;
                int rb = rowAdr;

                isCell = true;

                for (int k = 0; k < sheet.getNumMergedRegions(); k++) {
                    cellRangeAddress = sheet.getMergedRegion(k);
                    if (cellRangeAddress.getFirstColumn() == excelCell.getColumnIndex()
                            && cellRangeAddress.getFirstRow() == excelCell.getRowIndex()) {
                        cr = cellRangeAddress.getLastColumn() - refColAdr + 1;
                        rb = cellRangeAddress.getLastRow() - refRowAdr + 1;
                        break;
                    }

                    if (cellRangeAddress.getFirstColumn() <= excelCell.getColumnIndex()
                            && excelCell.getColumnIndex() <= cellRangeAddress.getLastColumn()
                            && cellRangeAddress.getFirstRow() <= excelCell.getRowIndex()
                            && excelCell.getRowIndex() <= cellRangeAddress.getLastRow()) {
                        isCell = false;
                    }
                }
                if (isCell) {
                    CCell cell = table.newCell();

                    cell.setCl(cl);
                    cell.setRt(rt);
                    cell.setCr(cr);
                    cell.setRb(rb);

                    fillCell(cell, excelCell);
                }
            }
        }

        this.rowIndex = endPnt.r + 1;

        // Обработка контекста таблицы
        /*
        CPoint namePnt = this.findPreviousPoint( this.sheet, TBL_NAME, refPnt.r - 1 );
        if ( null != namePnt )
        {
            row = sheet.getRow( namePnt.r);
            //excelCell = r.getCell( namePnt.c + 1 );
            excelCell = row.getCell( namePnt.c + 1, Row.CREATE_NULL_AS_BLANK );
            String name = extractCellValue( excelCell );
            //table.getContext().setName( name );
        }

        CPoint measurePnt = this.findPreviousPoint( this.sheet, TBL_MEASURE, refPnt.r - 1 );
        if ( null != measurePnt )
        {
            row = sheet.getRow( measurePnt.r);
            //excelCell = r.getCell( measurePnt.c + 1 );
            excelCell = row.getCell( measurePnt.c + 1, Row.CREATE_NULL_AS_BLANK );
            String measure = extractCellValue( excelCell );
            //table.getContext().setMeasure( measure );
        }
        */

        table.setSrcWorkbookFile(sourceWorkbookFile);
        table.setSrcSheetName(sheet.getSheetName());

        CellReference cellRef;
        cellRef = new CellReference(refPnt.r, refPnt.c);
        table.setSrcStartCellRef(cellRef.formatAsString());
        cellRef = new CellReference(endPnt.r, endPnt.c);
        table.setSrcEndCellRef(cellRef.formatAsString());

        return table;
    }

    private CPoint findPreviousPoint(Sheet sheet, String tag, int startRow) {
        for (int i = startRow; i > -1; i--) {
            Row row = sheet.getRow(i);
            if (null == row) continue;

            for (Cell cell : row) {
                String text = getFormatCellValue(cell);
                if (tag.equals(text))
                    return new CPoint(cell.getColumnIndex(), cell.getRowIndex());
            }
        }
        return null;
    }

    private CPoint findNextPoint(Sheet sheet, String tag, int startRow) {
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (null == row) continue;

            for (Cell cell : row) {
                String text = getFormatCellValue(cell);
                if (tag.equals(text))
                    return new CPoint(cell.getColumnIndex(), cell.getRowIndex());
            }
        }
        return null;
    }

    private String getFormatCellValue(Cell excelCell) {
        formulaEvaluator.evaluate(excelCell);
        return formatter.formatCellValue(excelCell, formulaEvaluator);
    }

    private CPoint findRefPoint(Sheet sheet, int startRow) {
        CPoint point = this.findNextPoint(sheet, REF_POINT_VAL, startRow);

        if (null != point) {
            point.c = point.c + 1;
            point.r = point.r + 1;
        }

        return point;
    }

    private CPoint findEndPoint(Sheet sheet, int startRow) {
        CPoint point = this.findNextPoint(sheet, END_POINT_VAL, startRow);

        if (null != point) {
            point.c = point.c - 1;
            point.r = point.r - 1;
        }

        return point;
    }

    private String extractCellFormulaValue(Cell excelCell) {
        String value = null;

        switch (excelCell.getCachedFormulaResultType()) {
            case Cell.CELL_TYPE_NUMERIC:
                value = Double.toString(excelCell.getNumericCellValue());
                break;

            case Cell.CELL_TYPE_STRING:
                value = excelCell.getStringCellValue();
                break;

            case Cell.CELL_TYPE_BOOLEAN:
                value = Boolean.toString(excelCell.getBooleanCellValue());
                break;
        }

        return value;
    }

    private String extractCellValue(Cell excelCell) {
        String value = null;

        switch (excelCell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(excelCell))
                    value = "DATE"; // TODO Какое-то странное значение
                else
                    value = Double.toString(excelCell.getNumericCellValue());
                break;

            case Cell.CELL_TYPE_STRING:
                value = excelCell.getRichStringCellValue().getString();
                break;

            case Cell.CELL_TYPE_BOOLEAN:
                value = Boolean.toString(excelCell.getBooleanCellValue());
                break;

            case Cell.CELL_TYPE_FORMULA:
                //value = excelCell.getCellFormula();
                value = extractCellFormulaValue(excelCell);
                break;

            case Cell.CELL_TYPE_BLANK:
                break;

            case Cell.CELL_TYPE_ERROR:
                break;
        }
        return value;
    }

    private DataFormatter formatter = new DataFormatter();
    private FormulaEvaluator formulaEvaluator;

    private boolean hasSuperscriptText(Cell excelCell) {
        if (excelCell.getCellType() != Cell.CELL_TYPE_STRING) return false;

        RichTextString richTextString = excelCell.getRichStringCellValue();
        if (null == richTextString) return false;

        int index = 0;
        int length = 0;
        XSSFFont font = null;

        XSSFRichTextString rts = (XSSFRichTextString) richTextString;
        XSSFWorkbook wb = (XSSFWorkbook) workbook;

        XSSFCellStyle style = ((XSSFCell) excelCell).getCellStyle();
        font = style.getFont();

        String richText = rts.getString();
        if (rts.numFormattingRuns() > 1) {
            for (int i = 0; i < rts.numFormattingRuns(); i++) {
                index = rts.getIndexOfFormattingRun(i);
                length = rts.getLengthOfFormattingRun(i);

                try {
                    font = rts.getFontOfFormattingRun(i);
                } catch (NullPointerException e) {
                    font = wb.getFontAt(XSSFFont.DEFAULT_CHARSET);
                    font.setTypeOffset(XSSFFont.SS_NONE);
                }

                String s = richText.substring(index, index + length);

                if (font.getTypeOffset() == XSSFFont.SS_SUPER)
                    return true;
            }
        } else {
            if (font.getTypeOffset() == XSSFFont.SS_SUPER)
                return true;
        }
        return false;
    }

    private String getNotSuperscriptText(Cell excelCell) {
        if (excelCell.getCellType() != Cell.CELL_TYPE_STRING) return null;
        RichTextString richTextString = excelCell.getRichStringCellValue();
        if (null == richTextString) return null;

        int index;
        int length;
        String text;

        XSSFRichTextString rts = (XSSFRichTextString) richTextString;
        XSSFWorkbook wb = (XSSFWorkbook) workbook;

        XSSFCellStyle style = ((XSSFCell) excelCell).getCellStyle();
        XSSFFont font = style.getFont();

        String richText = rts.getString();
        StringBuilder notSuperscriptRuns = new StringBuilder();
        if (rts.numFormattingRuns() > 1) {
            boolean wasNotSuperscriptRun = false;
            for (int i = 0; i < rts.numFormattingRuns(); i++) {
                index = rts.getIndexOfFormattingRun(i);
                length = rts.getLengthOfFormattingRun(i);

                try {
                    font = rts.getFontOfFormattingRun(i);
                } catch (NullPointerException e) {
                    font = wb.getFontAt(XSSFFont.DEFAULT_CHARSET);
                    font.setTypeOffset(XSSFFont.SS_NONE);
                }

                String s = richText.substring(index, index + length);

                if (font.getTypeOffset() == XSSFFont.SS_SUPER) {
                    if (wasNotSuperscriptRun) notSuperscriptRuns.append(" ");
                    wasNotSuperscriptRun = false;
                } else {
                    notSuperscriptRuns.append(s);
                    wasNotSuperscriptRun = true;
                }
            }
            text = notSuperscriptRuns.toString();
        } else {
            if (font.getTypeOffset() == XSSFFont.SS_SUPER)
                text = null;
            else
                text = richText;
        }
        return text;
    }

    private String getText(Cell excelCell) {
        String text = null;
        if (useCellValue)
            text = extractCellValue(excelCell);
        else
            text = getFormatCellValue(excelCell);
        return text;
    }

    private void fillCell(CCell cell, Cell excelCell) {
        String rawTextualContent = null;
        CellType cellType = null;

        String text = null;
        if (withoutSuperscript) {
            if (hasSuperscriptText(excelCell)) {
                text = getNotSuperscriptText(excelCell);
            } else {
                text = getText(excelCell);
            }
        } else {
            text = getText(excelCell);
        }
        cell.setText(text);

        rawTextualContent = getFormatCellValue(excelCell);
        cell.setRawText(rawTextualContent);

        switch (excelCell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(excelCell)) {
                    //rawTextualContent = "DATE"; // TODO Какое-то странное значение
                    cellType = CellType.DATE;
                } else {
                    cellType = CellType.NUMERIC;
                }
                break;

            case Cell.CELL_TYPE_STRING:
                cellType = CellType.STRING;
                break;

            case Cell.CELL_TYPE_BOOLEAN:
                cellType = CellType.BOOLEAN;
                break;

            case Cell.CELL_TYPE_FORMULA:
                cellType = CellType.FORMULA;
                break;

            case Cell.CELL_TYPE_BLANK:
                cellType = CellType.BLANK;
                break;

            case Cell.CELL_TYPE_ERROR:
                cellType = CellType.ERROR;
                break;
        }

        cell.setId(this.cellCount);

        cell.setType(cellType);

        int height = excelCell.getRow().getHeight();
        cell.setHeight(height);

        int width = excelCell.getSheet().getColumnWidth(excelCell.getColumnIndex());
        cell.setWidth(width);

        CellStyle excelCellStyle = excelCell.getCellStyle();
        CStyle cellStyle = cell.getStyle();
        //System.out.printf("cell = %s%n", cell.getText());
        fillCellStyle(cellStyle, excelCellStyle);

        String reference = new CellReference(excelCell).formatAsString();
        cell.setProvenance(reference);

        this.cellCount++;
    }

    private void fillCellStyle(CStyle cellStyle, CellStyle excelCellStyle) {
        Font excelFont = workbook.getFontAt(excelCellStyle.getFontIndex());
        // TODO надо переделать на CFont newFont(excelFont)
        //CFont font = new CFont();
        //cellStyle.setFont( font );
        CFont font = cellStyle.getFont();

        fillFont(font, excelFont);

        cellStyle.setHidden(excelCellStyle.getHidden());
        cellStyle.setLocked(excelCellStyle.getLocked());
        cellStyle.setWrapped(excelCellStyle.getWrapText());

        cellStyle.setIndention(excelCellStyle.getIndention());
        cellStyle.setRotation(excelCellStyle.getRotation());

        cellStyle.setHorzAlignment(this.getHorzAlignment(excelCellStyle.getAlignment()));
        cellStyle.setVertAlignment(this.getVertAlignment(excelCellStyle.getVerticalAlignment()));

        CBorder leftBorder = cellStyle.getLeftBorder();
        CBorder rightBorder = cellStyle.getRightBorder();
        CBorder topBorder = cellStyle.getTopBorder();
        CBorder bottomBorder = cellStyle.getBottomBorder();

        BorderType lbType = this.convertBorderType(excelCellStyle.getBorderLeft());
        BorderType rbType = this.convertBorderType(excelCellStyle.getBorderRight());
        BorderType tbType = this.convertBorderType(excelCellStyle.getBorderTop());
        BorderType bbType = this.convertBorderType(excelCellStyle.getBorderBottom());

        leftBorder.setType(lbType);
        rightBorder.setType(rbType);
        topBorder.setType(tbType);
        bottomBorder.setType(bbType);

        // Этот цвет "Fill Background Color" используется, только в тех случаях,
        // когда в ячейки используется узор. Тогда это цвет фона.
        // Без узора цвет фона задает "Fill Foreground Color"
        XSSFColor bgColor = (XSSFColor) excelCellStyle.getFillBackgroundColorColor();

        // Если Index цвета равен 64, то это значит, что ничего хорошего потом из такого цвета не получить,
        // это по сути тот же null для цвета
        if (null != bgColor && 64 != bgColor.getIndexed()) {
            String color = bgColor.getARGBHex();
            if (null != color) {
                color = color.substring(2);
                cellStyle.setBgColor(new CColor(color));
            }
        }

        // Этот цвет "Fill Background Color" задает цвет узора в тех случаях,
        // когда в ячейки используется узор. Без узора он задает цвет фона
        XSSFColor fgColor = (XSSFColor) excelCellStyle.getFillForegroundColorColor();

        if (null != fgColor && 64 != fgColor.getIndexed()) {
            String color = fgColor.getARGBHex();
            if (null != color) {
                color = color.substring(2);
                cellStyle.setFgColor(new CColor(color));
            }
        }

        // TODO Заполнить цвета границ
    }

    private CColor convertBorderColor(short originalExcelBorderColor) {
        // TODO получить цвет границы
        return new CColor("#000000");
    }

    // TODO конвертирует тип границы из CellStyle в CCellStyle
    private BorderType convertBorderType(short originalExcelBorderType) {
        if (originalExcelBorderType < 0 || originalExcelBorderType > 13)
            return null;

        switch (originalExcelBorderType) {
            case CellStyle.BORDER_NONE:
                return BorderType.NONE;
            case CellStyle.BORDER_THIN:
                return BorderType.THIN;
            case CellStyle.BORDER_MEDIUM:
                return BorderType.MEDIUM;
            case CellStyle.BORDER_DASHED:
                return BorderType.DASHED;
            case CellStyle.BORDER_HAIR:
                return BorderType.HAIR;
            case CellStyle.BORDER_THICK:
                return BorderType.THICK;
            case CellStyle.BORDER_DOUBLE:
                return BorderType.DOUBLE;
            case CellStyle.BORDER_DOTTED:
                return BorderType.DOTTED;
            case CellStyle.BORDER_MEDIUM_DASHED:
                return BorderType.MEDIUM_DASHED;
            case CellStyle.BORDER_DASH_DOT:
                return BorderType.DASH_DOT;
            case CellStyle.BORDER_MEDIUM_DASH_DOT:
                return BorderType.MEDIUM_DASH_DOT;
            case CellStyle.BORDER_DASH_DOT_DOT:
                return BorderType.DASH_DOT_DOT;
            case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
                return BorderType.MEDIUM_DASH_DOT_DOT;
            case CellStyle.BORDER_SLANTED_DASH_DOT:
                return BorderType.SLANTED_DASH_DOT;
            default:
                return null;
        }
    }

    private HorzAlignment getHorzAlignment(short originalExcelHorzAlignment) {
        if (originalExcelHorzAlignment < 0 || originalExcelHorzAlignment > 6)
            return null;

        switch (originalExcelHorzAlignment) {
            case CellStyle.ALIGN_GENERAL:
                return HorzAlignment.GENERAL;
            case CellStyle.ALIGN_LEFT:
                return HorzAlignment.LEFT;
            case CellStyle.ALIGN_CENTER:
                return HorzAlignment.CENTER;
            case CellStyle.ALIGN_RIGHT:
                return HorzAlignment.RIGHT;
            case CellStyle.ALIGN_FILL:
                return HorzAlignment.FILL;
            case CellStyle.ALIGN_JUSTIFY:
                return HorzAlignment.JUSTIFY;
            case CellStyle.ALIGN_CENTER_SELECTION:
                return HorzAlignment.CENTER_SELECTION;
            default:
                return null;
        }
    }

    private VertAlignment getVertAlignment(short originalExcelVertAlignment) {
        if (originalExcelVertAlignment < 0 || originalExcelVertAlignment > 3)
            return null;

        switch (originalExcelVertAlignment) {
            case CellStyle.VERTICAL_TOP:
                return VertAlignment.TOP;
            case CellStyle.VERTICAL_CENTER:
                return VertAlignment.CENTER;
            case CellStyle.VERTICAL_BOTTOM:
                return VertAlignment.BOTTOM;
            case CellStyle.VERTICAL_JUSTIFY:
                return VertAlignment.JUSTIFY;
            default:
                return null;
        }
    }

    private void fillFont(CFont font, Font excelFont) {
        font.setName(excelFont.getFontName());

        // TODO Задать цвет шрифта CFont font
        //font.setColor( excelFont.getColor() );

        font.setHeight(excelFont.getFontHeight());
        font.setHeightInPoints(excelFont.getFontHeightInPoints());

        // TODO Надо проверить значения Boldweight, сейчас все сделано наугад
        short boldWeight = excelFont.getBoldweight();
        if (boldWeight >= 700)
            font.setBold(true);

        font.setItalic(excelFont.getItalic());
        font.setStrikeout(excelFont.getStrikeout());

        byte underline = excelFont.getUnderline();
        if (underline != Font.U_NONE)
            font.setUnderline(true);
        if (underline == Font.U_DOUBLE || underline == Font.U_DOUBLE_ACCOUNTING)
            font.setDoubleUnderline(true);
    }

    private static final DataLoader INSTANCE = new DataLoader();

    public DataLoader() {
    }

    public static DataLoader getInstance() {
        return INSTANCE;
    }

    private static class CPoint {
        int c; // column index
        int r; // row index

        CPoint(int c, int r) {
            this.c = c;
            this.r = r;
        }
    }

    public  int[] getCellShift(){
        int[] shift = {cellShift.c, cellShift.r};
        return shift;
    }

    public Workbook getWorkbook() {
        return workbook;
    }
}
