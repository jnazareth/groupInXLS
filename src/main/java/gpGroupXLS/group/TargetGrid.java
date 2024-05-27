package gpGroupXLS.group;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.tabs.TabSummary;
import gpGroupXLS.xchg.ExchangeRateTable;
import gpGroupXLS.xchg.ExchangeRateTable.exchangePair;
import gpGroupXLS.xchg.ExchangeRateTable.targetCurrencies;
import gpGroupXLS.xls._Coordinates;

public class TargetGrid {
	private Set<Integer> cellInRange(int column, _Coordinates coords) {
		Set<Integer> sumCoords = coords.toCoordsSet() ;
		Set<Integer> sumColumn = new HashSet<Integer>() ;
		sumColumn.add(column);
		sumColumn.retainAll(sumCoords);
		return sumColumn ;
    }

	// [C13:C24] -> [C1:C12]		// adjusted by numberToSkip
	private _Coordinates adjustSumColumns(_Coordinates cd) {
		HashSet<Integer> newSumCoordsSet = new HashSet<Integer>() ;
		Set<Integer> sumCoordsSet = cd.toCoordsSet() ;
		for (Integer c : sumCoordsSet) {
			newSumCoordsSet.add(c - XLSProperties._numberToSkip + XLSProperties._groupNameColOffset);
		}
		_Coordinates newSumCoords = new _Coordinates(newSumCoordsSet) ;
		return newSumCoords;
	}

	private void buildExchangeGroupHeader(XSSFWorkbook wbSummary, String toCurrency) {
		XSSFSheet sheetSummary = wbSummary.getSheet(XLSProperties._SummarySheetName) ;
		if (sheetSummary == null) return;

		int lastRow = sheetSummary.getLastRowNum() + 2 ;
		Row newRow = sheetSummary.createRow(lastRow);
		if (newRow == null) return ;

		Cell cellTarget = newRow.createCell(0);
		cellTarget.setCellValue(toCurrency);
	}

	private void groupRows(XSSFSheet sheetSummary, int f, int t) {
		sheetSummary.groupRow(f-1, t);
	}

	private String makeSumFunction(int rF, int rT, String c) {
		return  "Sum(" + c + rF + ":" + c + rT + ")" ;
	}

	private void buildExchangeGroupSum(XSSFWorkbook wbSummary, String toCurrency, LinkedHashSet<Integer> rA, _Coordinates cd, String format) {
		XSSFSheet sheetSummary = wbSummary.getSheet(XLSProperties._SummarySheetName) ;
		if (sheetSummary == null) return ;

		int lastRow = sheetSummary.getLastRowNum();
		Row lRow = sheetSummary.getRow(lastRow);
		int lastColumn = lRow.getLastCellNum() ;

		Row newRow = sheetSummary.createRow(lastRow + 1);
		if (newRow == null) return ;

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		// _newFormatting 4/22
		_Coordinates newC = adjustSumColumns(cd) ;

		int rF = Collections.min(rA)+1;
		int rT = Collections.max(rA)+1;

		for (int col = 0; col < lastColumn; col++) {
			Cell cellTarget = newRow.createCell(col);
			//_newFormatting 4/22
			// col to add "total". Using 0 for this column ... should be via template
			if (col == XLSProperties._totalColumnPosition) {
				cellTarget.setCellValue(XLSProperties._Totals);
				continue ;
			}

			Set<Integer> sCol = cellInRange(col, newC);	//cd //_newFormatting 4/22
			if (sCol.size() != 0) {
				for (Integer c : sCol) {
					String cellR = CellReference.convertNumToColString(c);
					String formula = makeSumFunction(rF, rT, cellR) ;
					cellTarget.setCellFormula(formula) ;
					cellTarget.setCellStyle(cellStyle);	// apply format
				}
			}
		}
		groupRows(sheetSummary, rF, rT) ;
	}

	private int buildExchangeGroup2(XSSFWorkbook wbSummary, String toCurrency, int sourceRow, String rateReference, String format) {
		int rowAdded = -1 ;
		XSSFSheet sheetSummary = wbSummary.getSheet(XLSProperties._SummarySheetName) ;
		if (sheetSummary == null) return rowAdded;
		Row sRow = sheetSummary.getRow(sourceRow);
		if (sRow == null) return rowAdded;

		int lastRow = sheetSummary.getLastRowNum() + 1 ;
		Row newRow = sheetSummary.createRow(lastRow);
		rowAdded = newRow.getRowNum();

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		int lastColumn = sRow.getLastCellNum() ;
		for (int col = 0; col < lastColumn; col++) {
			Cell celldata = sRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
			if (celldata == null) {
				// The spreadsheet is empty in this cell
			} else {
				// Do something useful with the cell's contents
				Cell cellTarget = null;
				switch(celldata.getCellType()) {
					case STRING:
						String cV = celldata.getStringCellValue() ;
						cellTarget = newRow.createCell(col);
						cellTarget.setCellValue(cV);
						break;
					case NUMERIC:
						Double d = celldata.getNumericCellValue() ;
						cellTarget = newRow.createCell(col);

						CellReference cS = new CellReference(celldata) ;
						CellReference xR = new CellReference(rateReference);
						String sFormula = cS.formatAsString() + " * " +  xR.formatAsString() ;
						cellTarget.setCellFormula(sFormula) ;

						cellTarget.setCellStyle(cellStyle);	// apply format

						break;
					default:
						break;
				}
			}
		}
		return rowAdded;
	}

	public void buildTargetGrid(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			ExchangeRateTable ert = tg.m_xTable;
			LinkedHashMap<String, targetCurrencies> tGrid = ert.m_targetGrid;

			int r = 0 ;
			for (Map.Entry<String, targetCurrencies> tC : tGrid.entrySet()) {
				String toCurrency = tC.getKey();
				// add header
				buildExchangeGroupHeader(workBookGroup, toCurrency) ;

				// add converted grid
				LinkedHashSet<Integer> rowsAdded = new LinkedHashSet<Integer>() ;
				targetCurrencies tCs = tC.getValue();
				String cFormat = tCs.m_CurrencyInfo.getCurrencyFormat(toCurrency);
				ArrayList<exchangePair> tR = tCs.m_targetRates;
				for (int i = 0; i < tR.size(); i++) {
					String fromCurrency = tR.get(i).fromCurrency ;
					String rateReference = ert.getRateReference(i, r) ;

					TabSummary ts = tg.m_tabSummary ;
					int row = ts.m_groupTabs.get(i).rowNumber;
					if (row != -1) {
						int rA = buildExchangeGroup2(workBookGroup, toCurrency, row, rateReference, cFormat) ;
						if (rA != -1) rowsAdded.add(rA) ;
					}
				}

				// add sum
				_Coordinates cd = tg.m_tabSummary.getCoords() ;
				buildExchangeGroupSum(workBookGroup, toCurrency, rowsAdded, cd, cFormat) ;
				r++ ;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
