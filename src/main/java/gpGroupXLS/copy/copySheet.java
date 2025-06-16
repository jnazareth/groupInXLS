package gpGroupXLS.copy;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.group.XLSProperties;
import gpGroupXLS.group.TabGroup;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.tabs.TabSummary2.TabEntry2;
import gpGroupXLS.tabs.TabSummary2.TabGroupBase;;

public class CopySheet {

    public CopySheet() {
        // constructor
    }

    private void copyCellStyle(XSSFWorkbook targetWorkbook, XSSFWorkbook sourceWorkbook, Cell sourceCell, Cell targetCell) {
        XSSFCellStyle targetCellStyle = targetWorkbook.createCellStyle();
        DataFormat dataFormat = targetWorkbook.createDataFormat();
        String formatString = sourceCell.getCellStyle().getDataFormatString();
        targetCellStyle.setDataFormat(dataFormat.getFormat(formatString));

        XSSFFont sourceFont = sourceWorkbook.getFontAt(sourceCell.getCellStyle().getFontIndex());
        XSSFFont targetFont = targetWorkbook.createFont();
        copyFontProperties(sourceFont, targetFont);
        targetCellStyle.setFont(targetFont);

        copyCellStyleProperties(sourceCell, targetCellStyle);
        targetCell.setCellStyle(targetCellStyle);
    }

    private void copyFontProperties(XSSFFont sourceFont, XSSFFont targetFont) {
        targetFont.setBold(sourceFont.getBold());
        targetFont.setColor(sourceFont.getColor());
        targetFont.setFontHeight(sourceFont.getFontHeight());
        targetFont.setFontName(sourceFont.getFontName());
        targetFont.setItalic(sourceFont.getItalic());
        targetFont.setStrikeout(sourceFont.getStrikeout());
        targetFont.setTypeOffset(sourceFont.getTypeOffset());
        targetFont.setUnderline(sourceFont.getUnderline());
        targetFont.setCharSet(sourceFont.getCharSet());
    }

    private void copyCellStyleProperties(Cell sourceCell, XSSFCellStyle targetCellStyle) {
        targetCellStyle.setAlignment(sourceCell.getCellStyle().getAlignment());
        targetCellStyle.setHidden(sourceCell.getCellStyle().getHidden());
        targetCellStyle.setLocked(sourceCell.getCellStyle().getLocked());
        targetCellStyle.setWrapText(sourceCell.getCellStyle().getWrapText());
        targetCellStyle.setBorderBottom(sourceCell.getCellStyle().getBorderBottom());
        targetCellStyle.setBorderLeft(sourceCell.getCellStyle().getBorderLeft());
        targetCellStyle.setBorderRight(sourceCell.getCellStyle().getBorderRight());
        targetCellStyle.setBorderTop(sourceCell.getCellStyle().getBorderTop());
        targetCellStyle.setFillBackgroundColor(sourceCell.getCellStyle().getFillBackgroundColor());
        targetCellStyle.setFillForegroundColor(sourceCell.getCellStyle().getFillForegroundColor());
        targetCellStyle.setFillPattern(sourceCell.getCellStyle().getFillPattern());
        targetCellStyle.setIndention(sourceCell.getCellStyle().getIndention());
        targetCellStyle.setBottomBorderColor(sourceCell.getCellStyle().getBottomBorderColor());
        targetCellStyle.setLeftBorderColor(sourceCell.getCellStyle().getLeftBorderColor());
        targetCellStyle.setRightBorderColor(sourceCell.getCellStyle().getRightBorderColor());
        targetCellStyle.setTopBorderColor(sourceCell.getCellStyle().getTopBorderColor());
        targetCellStyle.setRotation(sourceCell.getCellStyle().getRotation());
        targetCellStyle.setVerticalAlignment(sourceCell.getCellStyle().getVerticalAlignment());
    }

    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            default:
                break;
        }
    }

    private int copySheetContents(XSSFWorkbook targetWorkbook, XSSFSheet targetSheet, XSSFSheet sourceSheet) {
        int lastRowNum = sourceSheet.getLastRowNum();
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row sourceRow = sourceSheet.getRow(rowIndex);
            if (sourceRow != null) {
                Row targetRow = targetSheet.createRow(rowIndex);
                copyRowContents(targetWorkbook, sourceSheet.getWorkbook(), sourceRow, targetRow);
            }
        }
        return lastRowNum;
    }

    private void copyRowContents(XSSFWorkbook targetWorkbook, XSSFWorkbook sourceWorkbook, Row sourceRow, Row targetRow) {
        int lastCellNum = sourceRow.getLastCellNum();
        for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
            Cell sourceCell = sourceRow.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(cellIndex);
                copyCellStyle(targetWorkbook, sourceWorkbook, sourceCell, targetCell);
                copyCellValue(sourceCell, targetCell);
            }
        }
    }

    private String generateSheetName(String fileName, String groupName) {
        final String separator = ".";
        String prefix = fileName.contains(separator) ? fileName.substring(0, fileName.indexOf(separator) + 1) : "";
        String suffix = groupName.contains(separator) ? groupName.substring(groupName.lastIndexOf(separator) + 1) : "";
        int hash = Math.abs((fileName + separator + groupName).hashCode() % 10000);
        return prefix + "xlsx." + suffix + separator + hash;
    }

    public void buildCopySheets(XSSFWorkbook targetWorkbook, TabGroup tabGroup) {
        TabSummary2 tabSummary = tabGroup.getTabSummary();
        for (TabGroupBase tabGroupBase : tabSummary.getGroupTabs()) {
            TabEntry2 groupItem = tabGroupBase.getTabEntry();
            try (FileInputStream fileInputStream = new FileInputStream(new File(groupItem.getFileName()));
                 XSSFWorkbook sourceWorkbook = new XSSFWorkbook(fileInputStream)) {

                XSSFCreationHelper creationHelper = targetWorkbook.getCreationHelper();
                XSSFHyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.DOCUMENT);

                String sheetName = copySheet(sourceWorkbook, groupItem.getFileName(), groupItem.getGroupName(), targetWorkbook);
                if (sheetName != null) {
                    updateSummarySheet(targetWorkbook, tabGroupBase.getRowNumber(), groupItem.getGroupName(), sheetName, hyperlink);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void updateSummarySheet(XSSFWorkbook targetWorkbook, int rowNumber, String groupName, String sheetName, XSSFHyperlink hyperlink) {
        if (rowNumber != -1) {
            XSSFSheet summarySheet = targetWorkbook.getSheet(XLSProperties.SUMMARY_SHEET_NAME);
            if (summarySheet != null) {
                Row row = summarySheet.getRow(rowNumber);
                if (row != null) {
                    Cell cell = row.createCell(0);
                    String cellReference = sheetName + "!$A$1";
                    CellReference cellRef = new CellReference(cellReference);
                    cell.setCellValue(groupName);
                    hyperlink.setAddress(cellRef.formatAsString());
                    cell.setHyperlink(hyperlink);
                }
            }
        }
    }

    public String copySheet(XSSFWorkbook sourceWorkbook, String fileName, String groupName, XSSFWorkbook targetWorkbook) {
        XSSFSheet sourceSheet = sourceWorkbook.getSheet(groupName);
        if (sourceSheet == null) return null;

        String targetSheetName = generateSheetName(fileName, groupName);
        XSSFSheet targetSheet = targetWorkbook.createSheet(targetSheetName);
        if (targetSheet == null) return null;

        copySheetContents(targetWorkbook, targetSheet, sourceSheet);
        return targetSheetName;
    }

    @Override
    public String toString() {
        return super.toString();
    }
}