package gpGroupXLS.group;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.xchg.ExchangeRateTable;
import gpGroupXLS.xchg.ExchangeRateTable.ExchangePair;
import gpGroupXLS.xchg.ExchangeRateTable.TargetCurrencies;

public class XRateGrid {

    private CellReference addExchangeRateSheet(XSSFWorkbook workbook, String[][] rateArray) {
        try {
            XSSFSheet sheet = workbook.getSheet(XLSProperties.XRATES_SHEET_NAME);
            if (sheet == null) {
                sheet = workbook.createSheet(XLSProperties.XRATES_SHEET_NAME);
            }

            for (int i = 0; i < rateArray.length; i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < rateArray[i].length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(rateArray[i][j]);
                }
            }

            return new CellReference(sheet.getSheetName() + "!A1");
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private void populateExchangeRates(String[][] rateGrid, ArrayList<String> columnData, int index) {
        if (rateGrid.length != columnData.size()) return;
        try {
            String[] values = columnData.toArray(new String[0]); 
            for (int i = 0; i < rateGrid.length; i++) {
                rateGrid[i][index] = values[i];
            }
            columnData.clear();
        } catch (ArrayIndexOutOfBoundsException e) {
            System.err.println("populateExchangeRates::Exception::" + e.getMessage());
        }
    }

    private String[][] buildRateArrayGrid(TabGroup tabGroup) {
        try {
            final String fromToHeader = "from (C) | to (R)";
            final String dateHeader = "date";

            ExchangeRateTable exchangeRateTable = tabGroup.getExchangeRateTable();
            LinkedHashMap<String, TargetCurrencies> targetGrid = exchangeRateTable.getTargetGrid();
            int numCurrencies = targetGrid.size();
            int numRates = targetGrid.values().iterator().next().getTargetRates().size();

            int headerRows = 1;
            int columns = (numCurrencies * 2) + headerRows;
            int rows = numRates + headerRows;
            String[][] rateGrid = new String[rows][columns];

            ArrayList<String> fromCurrencies = new ArrayList<>(columns);
            ArrayList<String> toCurrencies = new ArrayList<>(columns);
            ArrayList<String> dates = new ArrayList<>(columns);

            boolean isFirstColumn = true;
            boolean isColumnheader = true;
            int columnIndex = 0;
            for (Map.Entry<String, TargetCurrencies> entry : targetGrid.entrySet()) {
                TargetCurrencies currencies = entry.getValue();
                for (ExchangePair pair : currencies.getTargetRates()) {
                    if  (columnIndex == 0) {
                            if (isFirstColumn) {
                                fromCurrencies.add(fromToHeader);
                                isFirstColumn = false;
                            }
                            fromCurrencies.add(pair.getFromCurrency());
                    }
                    if (isColumnheader) {
                        toCurrencies.add(pair.getToCurrency());
                        dates.add(dateHeader);
                        isColumnheader = false;
                    }
                    toCurrencies.add(String.valueOf(pair.getRateDate().getRate()));
                    dates.add(pair.getRateDate().getDate());
                }
                isColumnheader = true;
                if (columnIndex == 0) {
                    populateExchangeRates(rateGrid, fromCurrencies, columnIndex);
                }
                populateExchangeRates(rateGrid, toCurrencies, columnIndex + 1);
                populateExchangeRates(rateGrid, dates, columnIndex + 2);

                columnIndex += 2;
            }
            return rateGrid;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private CellReference addExchangeRateSheet(XSSFWorkbook workbook, TabGroup tabGroup) {
        try {
            String[][] rateArray = buildRateArrayGrid(tabGroup);
            if (rateArray == null) return null;

            return addExchangeRateSheet(workbook, rateArray);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public void buildExchangeRateGrid(XSSFWorkbook workbook, TabGroup tabGroup) {
        try {
            CellReference cellReference = addExchangeRateSheet(workbook, tabGroup);
            if (cellReference == null) return;

            ExchangeRateTable exchangeRateTable = tabGroup.getExchangeRateTable();
            exchangeRateTable.buildRateReferenceGrid(cellReference);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String toString() {
        return "XRateGrid{}";
    }
}