package gpGroupXLS.group;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.tabs.TabSummary2.TabEntry2;
import gpGroupXLS.tabs.TabSummary2.TabGroupBase;;

public class BaseGrid {

    private int locateSourceRow(XSSFSheet sheet) {
        final String pivotStartIndicator1 = "Values";
        final String pivotStartIndicator2 = "Row Labels";
        int sourceRow = -1;

        try {
            int lastRow = sheet.getLastRowNum();
            if (lastRow == -1) return sourceRow;

            boolean pivotTableStartFound = false;
            for (int r = lastRow; r > 0; r--) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int col = 0; col < row.getLastCellNum(); col++) {
                        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            if (cellValue.equalsIgnoreCase(pivotStartIndicator1) || cellValue.equalsIgnoreCase(pivotStartIndicator2)) {
                                pivotTableStartFound = true;
                                break;
                            }
                            if (pivotTableStartFound && !cellValue.isEmpty()) {
                                return r;
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("locateSourceRow::Exception::" + e.getMessage());
        }
        return sourceRow;
    }

    private int copyContents(XSSFWorkbook summaryWorkbook, XSSFSheet summarySheet, String groupName, XSSFSheet sourceSheet, String format, LinkedHashSet<Integer> rowsToCopy) {
        int rowAdded = -1;
        int currentRowNum = summarySheet.getLastRowNum() + 1;

        int sourceRow = locateSourceRow(sourceSheet);
        if (sourceRow != -1) rowsToCopy.add(sourceRow);

        XSSFCellStyle cellStyle = createCellStyle(summaryWorkbook, format);

        for (Integer sourceRowIndex : rowsToCopy) {
            Row sourceRowObj = sourceSheet.getRow(sourceRowIndex);
            if (sourceRowObj != null) {
                Row targetRow = summarySheet.createRow(currentRowNum++);
                rowAdded = targetRow.getRowNum();
                copyRowData(sourceRowObj, targetRow, cellStyle, groupName);
            }
        }
        return rowAdded;
    }

    private XSSFCellStyle createCellStyle(XSSFWorkbook workbook, String format) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));
        return cellStyle;
    }

    private void copyRowData(Row sourceRow, Row targetRow, XSSFCellStyle cellStyle, String groupName) {
        int colInsertPosition = 0;
        for (int col = XLSProperties.numberToSkip - XLSProperties.GROUP_NAME_COLUMN_OFFSET; col < sourceRow.getLastCellNum(); col++) {
            if (col == XLSProperties.numberToSkip - XLSProperties.GROUP_NAME_COLUMN_OFFSET) {
                Cell groupNameCell = targetRow.createCell(colInsertPosition++);
                groupNameCell.setCellValue(groupName);
                continue;
            }

            Cell sourceCell = sourceRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(colInsertPosition++);
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        targetCell.setCellStyle(cellStyle);
                        break;
                    default:
                        break;
                }
            }
        }
    }

    private int extractSheetData(XSSFWorkbook inputWorkbook, String groupName, XSSFWorkbook groupWorkbook, String format, boolean isFirst) {
        XSSFSheet sourceSheet = inputWorkbook.getSheet(groupName);
        if (sourceSheet == null) return -1;

        XSSFSheet summarySheet = groupWorkbook.getSheet(XLSProperties.SUMMARY_SHEET_NAME);
        if (summarySheet == null) summarySheet = groupWorkbook.createSheet(XLSProperties.SUMMARY_SHEET_NAME);

        LinkedHashSet<Integer> rowsToCopy = new LinkedHashSet<>();
        if (isFirst) {
            rowsToCopy.add(0); // header# 1 (row 0)
            rowsToCopy.add(1); // header# 2 (row 1)
        }
        return copyContents(groupWorkbook, summarySheet, groupName, sourceSheet, format, rowsToCopy);
    }

    public void buildBaseGrid(XSSFWorkbook groupWorkbook, TabGroup tabGroup) {
        try {
            boolean isFirst = true;
            for (TabGroupBase tabGroupBase : tabGroup.getTabSummary().getGroupTabs()) {
                TabEntry2 groupItem = tabGroupBase.getTabEntry();
                try (FileInputStream fileInputStream = new FileInputStream(new File(groupItem.getFileName()));
                     XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInputStream)) {

                    int rowInserted = extractSheetData(inputWorkbook, groupItem.getGroupName(), groupWorkbook, groupItem.getFormat(), isFirst);
                    if (rowInserted != -1) tabGroupBase.setRowNumber(rowInserted);
                    isFirst = false;
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}