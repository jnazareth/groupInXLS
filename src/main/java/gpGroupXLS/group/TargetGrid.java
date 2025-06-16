package gpGroupXLS.group;

import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.xchg.ExchangeRateTable;
import gpGroupXLS.xchg.ExchangeRateTable.ExchangePair;
import gpGroupXLS.xchg.ExchangeRateTable.TargetCurrencies;
import gpGroupXLS.xls._Coordinates;

public class TargetGrid {

    private Set<Integer> getCellsInRange(int column, _Coordinates coords) {
        Set<Integer> sumCoords = coords.toCoordsSet();
        Set<Integer> sumColumn = new HashSet<>();
        sumColumn.add(column);
        sumColumn.retainAll(sumCoords);
        return sumColumn;
    }

    private _Coordinates adjustSumColumns(_Coordinates coords) {
        HashSet<Integer> adjustedCoordsSet = new HashSet<>();
        for (Integer c : coords.toCoordsSet()) {
            adjustedCoordsSet.add(c - XLSProperties.numberToSkip + XLSProperties.GROUP_NAME_COLUMN_OFFSET);
        }
        return new _Coordinates(adjustedCoordsSet);
    }

    private void addExchangeGroupHeader(XSSFWorkbook workbook, String toCurrency) {
        XSSFSheet sheet = workbook.getSheet(XLSProperties.SUMMARY_SHEET_NAME);
        if (sheet == null) return;

        int lastRow = sheet.getLastRowNum() + 2;
        Row newRow = sheet.createRow(lastRow);
        if (newRow == null) return;

        Cell cell = newRow.createCell(0);
        cell.setCellValue(toCurrency);
    }

    private void groupRows(XSSFSheet sheet, int from, int to) {
        sheet.groupRow(from - 1, to);
    }

    private String createSumFormula(int fromRow, int toRow, String column) {
        return "SUM(" + column + fromRow + ":" + column + toRow + ")";
    }

    private void addExchangeGroupSum(XSSFWorkbook workbook, String toCurrency, LinkedHashSet<Integer> rowIndexes, _Coordinates coords, String format) {
        XSSFSheet sheet = workbook.getSheet(XLSProperties.SUMMARY_SHEET_NAME);
        if (sheet == null) return;

        int lastRow = sheet.getLastRowNum();
        Row lastRowObj = sheet.getRow(lastRow);
        int lastColumn = lastRowObj.getLastCellNum();

        Row newRow = sheet.createRow(lastRow + 1);
        if (newRow == null) return;

        XSSFCellStyle cellStyle = createCellStyle(workbook, format);
        _Coordinates adjustedCoords = adjustSumColumns(coords);

        int fromRow = Collections.min(rowIndexes) + 1;
        int toRow = Collections.max(rowIndexes) + 1;

        for (int col = 0; col < lastColumn; col++) {
            Cell cell = newRow.createCell(col);
            if (col == XLSProperties.TOTAL_COLUMN_POSITION) {
                cell.setCellValue(XLSProperties.TOTALS_LABEL);
                continue;
            }

            Set<Integer> columnsInRange = getCellsInRange(col, adjustedCoords);
            if (!columnsInRange.isEmpty()) {
                for (Integer column : columnsInRange) {
                    String columnRef = CellReference.convertNumToColString(column);
                    String formula = createSumFormula(fromRow, toRow, columnRef);
                    cell.setCellFormula(formula);
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        groupRows(sheet, fromRow, toRow);
    }

    private XSSFCellStyle createCellStyle(XSSFWorkbook workbook, String format) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));
        return cellStyle;
    }

    private int addExchangeGroup(XSSFWorkbook workbook, String toCurrency, int sourceRow, String rateReference, String format) {
        XSSFSheet sheet = workbook.getSheet(XLSProperties.SUMMARY_SHEET_NAME);
        if (sheet == null) return -1;

        Row sourceRowObj = sheet.getRow(sourceRow);
        if (sourceRowObj == null) return -1;

        int lastRow = sheet.getLastRowNum() + 1;
        Row newRow = sheet.createRow(lastRow);

        XSSFCellStyle cellStyle = createCellStyle(workbook, format);

        int lastColumn = sourceRowObj.getLastCellNum();
        for (int col = 0; col < lastColumn; col++) {
            Cell sourceCell = sourceRowObj.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (sourceCell != null) {
                Cell targetCell = newRow.createCell(col);
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        String formula = createFormula(sourceCell, rateReference);
                        targetCell.setCellFormula(formula);
                        targetCell.setCellStyle(cellStyle);
                        break;
                    default:
                        break;
                }
            }
        }
        return newRow.getRowNum();
    }

    private String createFormula(Cell sourceCell, String rateReference) {
        CellReference sourceRef = new CellReference(sourceCell);
        CellReference rateRef = new CellReference(rateReference);
        return sourceRef.formatAsString() + " * " + rateRef.formatAsString();
    }

    public void buildTargetGrid(XSSFWorkbook workbook, TabGroup tabGroup) {
        try {
            ExchangeRateTable exchangeRateTable = tabGroup.getExchangeRateTable();
            LinkedHashMap<String, TargetCurrencies> targetGrid = exchangeRateTable.getTargetGrid();

            int rateIndex = 0;
            for (Map.Entry<String, TargetCurrencies> entry : targetGrid.entrySet()) {
                String toCurrency = entry.getKey();
                addExchangeGroupHeader(workbook, toCurrency);

                LinkedHashSet<Integer> rowsAdded = new LinkedHashSet<>();
                TargetCurrencies currencies = entry.getValue();
                String currencyFormat = currencies.getCurrencyInfo().getCurrencyFormat(toCurrency);
                List<ExchangePair> targetRates = currencies.getTargetRates();

                for (int i = 0; i < targetRates.size(); i++) {
                    String rateReference = exchangeRateTable.getRateReference(i, rateIndex);
                    int rowNumber = tabGroup.getTabSummary().getGroupTabs().get(i).getRowNumber();
                    if (rowNumber != -1) {
                        int addedRow = addExchangeGroup(workbook, toCurrency, rowNumber, rateReference, currencyFormat);
                        if (addedRow != -1) rowsAdded.add(addedRow);
                    }
                }

                _Coordinates coords = tabGroup.getTabSummary().getCoords();
                addExchangeGroupSum(workbook, toCurrency, rowsAdded, coords, currencyFormat);
                rateIndex++;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String toString() {
        return "TargetGrid{}";
    }
}