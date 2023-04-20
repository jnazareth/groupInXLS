package gpGroupXLS.group;

import gpGroupXLS.json.ReadJson;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.tabs.TabSummary2.tabEntry2;
import gpGroupXLS.tabs.TabSummary2.tabGroupBase;
import gpGroupXLS.xchg.ExchangeRateTable2;
import gpGroupXLS.xchg.ExchangeRateTable2.exchangePair;
import gpGroupXLS.xchg.ExchangeRateTable2.targetCurrencies;
import gpGroupXLS.xls._Coordinates;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.XSSFTable;

public class groupXLS2 {
	private final String _SHEETNAME = "Summary" ;

	public File InitializeXLS(String fName) {
        File f = null ;
		try {
            String xlsFile = fName;

			f = new File(xlsFile);
            String filetoRecreate = f.getName();
            boolean dFile = false;
			if (f.exists()) {
                dFile = f.delete();
				if (!dFile) System.out.println("failed to delete:" + filetoRecreate);
            }
			File f2 = new File(filetoRecreate);
			f = f2;
			XSSFWorkbook workBook = new XSSFWorkbook();
			try (FileOutputStream fileOut = new FileOutputStream(f)) {
				workBook.write(fileOut);
				workBook.close();
				fileOut.close() ;
			}
			return f;
        } catch (FileNotFoundException fnf) {
        } catch (IOException ioe) {
        }
        return f;
	}

	private Set<Integer> cellInRange(int column, _Coordinates coords) {
		Set<Integer> sumCoords = coords.toCoordsSet() ;
		Set<Integer> sumColumn = new HashSet<Integer>() ;
		sumColumn.add(column);
		sumColumn.retainAll(sumCoords);
		return sumColumn ;
    }

	private void buildExchangeGroupHeader(XSSFWorkbook wbSummary, String toCurrency) {
		XSSFSheet sheetSummary = wbSummary.getSheet(_SHEETNAME) ;
		if (sheetSummary == null) return;

		int lastRow = sheetSummary.getLastRowNum() + 2 ;
		Row newRow = sheetSummary.createRow(lastRow);
		if (newRow == null) return ;

		Cell cellTarget = newRow.createCell(0);
		cellTarget.setCellValue(toCurrency);
	}

	private String makeSumFunction(int rF, int rT, String c) {
		String f =  "Sum(" + c + rF + ":" + c + rT + ")" ;
		return f;
	}

	private void buildExchangeGroupSum(XSSFWorkbook wbSummary, String toCurrency, LinkedHashSet<Integer> rA, _Coordinates cd, String format) {
		XSSFSheet sheetSummary = wbSummary.getSheet(_SHEETNAME) ;
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

		int rF = Collections.min(rA)+1;
		int rT = Collections.max(rA)+1;

		for (int col = 0; col < lastColumn; col++) {
			Cell cellTarget = newRow.createCell(col);
			Set<Integer> sCol = cellInRange(col, cd);
			if (sCol.size() != 0) {
				for (Integer c : sCol) {
					String cellR = CellReference.convertNumToColString(c);
					String formula = makeSumFunction(rF, rT, cellR) ;
					cellTarget.setCellFormula(formula) ;
					cellTarget.setCellStyle(cellStyle);	// apply format
				}
			}
		}
	}

	private int buildExchangeGroup2(XSSFWorkbook wbSummary, String toCurrency, int sourceRow, Double rate, String format) {
		int rowAdded = -1 ;
		XSSFSheet sheetSummary = wbSummary.getSheet(_SHEETNAME) ;
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
						cellTarget.setCellValue(d * rate);	// apply rate
						cellTarget.setCellStyle(cellStyle);	// apply format
						//CellAddress c1 = cellTarget.getAddress();
						//System.out.println("cell address:" + c1.formatAsString() + "|c:" + c1.getColumn() + "|r:" + c1.getRow());
						break;
					default:
						break;
				}
			}
		}
		return rowAdded;
	}

	private int locateSourceRow(XSSFSheet sheet) {
		int sourceRow = -1;
		try {
			final String pivotStartIndicator1 = "Values" ;
			final String pivotStartIndicator2 = "Row Labels" ;

			int lastRow = sheet.getLastRowNum() ;
			if (lastRow == -1) return sourceRow ;

			boolean bPivotTableStartFound = false ;
			for (int r = lastRow; r > 0; r--) {
				Row aRow = sheet.getRow(r) ;
				if (aRow != null) {
					boolean skipColumns = false;
					int lastColumn = aRow.getLastCellNum() ;
					for (int col = 0; ((col < lastColumn) && (!skipColumns)); col++) {
						Cell celldata = aRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						if (celldata == null) {
							// The spreadsheet is empty in this cell
						} else {
							switch(celldata.getCellType()) {
								case STRING:
									String cV = celldata.getStringCellValue() ;
									if ((cV.compareToIgnoreCase(pivotStartIndicator1) == 0) || (cV.compareToIgnoreCase(pivotStartIndicator2) == 0)) {
										bPivotTableStartFound = true ;
										skipColumns = true;
										continue ;
									}
									if ((bPivotTableStartFound) && (cV.length() != 0)) {
										sourceRow = r ;
										return sourceRow;
									}
									break;
								default:
									break;
							}
						}
					}
				}
			}
			return sourceRow ;
		} catch (Exception e) {
			System.err.println("locateSourceRow::Exception::" + e.getMessage()) ;
			return sourceRow;
		}
	}

	private int copyContents3(XSSFWorkbook wbSummary, XSSFSheet sheetSummary, XSSFSheet sheetToBeGrouped, String format, LinkedHashSet<Integer> rowsToCopy) {
		int rowAdded = -1 ;
		int row = sheetSummary.getLastRowNum() + 1 ;

		// does not work. Table not created using POI API
		//readTables(sheetToBeGrouped) ;

		// locate table. strange way of locating table position.
		// POI API only identifies tables if created with SDK
		// unable to locate pivot either (previous attempts failed)
		// resorting to parsing the sheet using poivot header column as "anchor". Very clumsy.
		int sourceRow = locateSourceRow(sheetToBeGrouped);
		//System.out.println("source row for " + sheetToBeGrouped.getSheetName() + ":" + sourceRow);

		LinkedHashSet<Integer> rowsToCopy2 = new LinkedHashSet<Integer>();
		rowsToCopy2 = (LinkedHashSet)rowsToCopy.clone();
		if (sourceRow != -1) rowsToCopy2.add(sourceRow) ;
		Integer[] sourceRows = new Integer [rowsToCopy2.size()];
		sourceRows = rowsToCopy2.toArray(sourceRows) ;

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		Integer rNum = -1 ;
		for (int r = 0; r < sourceRows.length; r++) {
			Row rowitr = sheetToBeGrouped.getRow(sourceRows[r]) ;
			if (rowitr != null) {
				Row currentRow = sheetSummary.createRow(row);
				rowAdded = currentRow.getRowNum();

				int lastColumn = rowitr.getLastCellNum() ;
				for (int col = 0; col < lastColumn; col++) {
					Cell celldata = rowitr.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					if (celldata == null) {
						// The spreadsheet is empty in this cell
					} else {
						Cell cellTarget = null;
						switch(celldata.getCellType()) {
							case STRING:
								String cV = celldata.getStringCellValue() ;
								cellTarget = currentRow.createCell(col);
								cellTarget.setCellValue(cV);
								break;
							case NUMERIC:
								Double d = celldata.getNumericCellValue() ;
								cellTarget = currentRow.createCell(col);
								cellTarget.setCellValue(d);
								cellTarget.setCellStyle(cellStyle);
								break;
							default:
								break;
						}
					}
				}
			}
			row = sheetSummary.getLastRowNum() + 1 ;
		}
		return (rowAdded) ;
	}

	private int extractSheetData(XSSFWorkbook workBookIn, String groupName, XSSFWorkbook workBookGroup, String format, boolean bFirst) {
		XSSFSheet sheetToBeGrouped = workBookIn.getSheet(groupName) ;
		if (sheetToBeGrouped == null) return -1 ;

		XSSFSheet sheetSummary = workBookGroup.getSheet(_SHEETNAME) ;
		if (sheetSummary == null) sheetSummary = workBookGroup.createSheet(_SHEETNAME);

		LinkedHashSet<Integer> rowsToCopy = new LinkedHashSet<Integer>();
		if (bFirst) {
			rowsToCopy.add(0) ;		// header# 1 (row 0)
			rowsToCopy.add(1) ;		// header# 2 (row 1)
		}
		int rInserted = copyContents3(workBookGroup, sheetSummary, sheetToBeGrouped, format, rowsToCopy) ;
		return rInserted;
	}

	private void buildBaseGrid(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			boolean bFirst = true ;
			XSSFWorkbook workBookIn = null;
			TabSummary2 ts2 = tg.m_tabSummary ;
			for (tabGroupBase tgb : ts2.m_groupTabs) {
				tabEntry2 gItem = tgb.te;
				File fIn = new File(gItem.fileName);
				FileInputStream fileIn = new FileInputStream(fIn);
				workBookIn = new XSSFWorkbook(fileIn);
				int r = extractSheetData(workBookIn, gItem.groupName, workBookGroup, gItem.format, bFirst) ;
				if (r != -1) tgb.rowNumber = r ;
				fileIn.close() ;
				bFirst = false ;
			}
			try {
				workBookIn.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void buildTargetGrid(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			ExchangeRateTable2 ert2 = tg.m_xTable;
			LinkedHashMap<String, targetCurrencies> tGrid = ert2.m_targetGrid;
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
					Double rate = tR.get(i).rate;
					TabSummary2 ts2 = tg.m_tabSummary ;
					int row = ts2.m_groupTabs.get(i).rowNumber;
					if (row != -1) {
						int rA = buildExchangeGroup2(workBookGroup, toCurrency, row, rate, cFormat) ;
						if (rA != -1) rowsAdded.add(rA) ;
					}
				}

				// add sum
				_Coordinates cd = tg.m_tabSummary.getCoords() ;
				buildExchangeGroupSum(workBookGroup, toCurrency, rowsAdded, cd, cFormat) ;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private boolean buildGroupXLS2(File hXLS, tabGroup tg) {
		try {
			File fGroup = new File(hXLS.getName());
			FileInputStream fileGroup = new FileInputStream(fGroup);
			XSSFWorkbook workBookGroup = new XSSFWorkbook(fileGroup);

			buildBaseGrid(workBookGroup, tg) ;
			buildTargetGrid(workBookGroup, tg) ;

			try {
				FileOutputStream outputStream = new FileOutputStream(hXLS.getName());
				workBookGroup.write(outputStream);
				workBookGroup.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			fileGroup.close() ;

			return true ;
		} catch (FileNotFoundException nfe) {
			nfe.printStackTrace();
			return false ;
		} catch (Exception e) {
			e.printStackTrace();
			return false ;
		}
	}

	private tabGroup readFromJSON(String configJSON, String xlsGroupFile) {
		try {
			ReadJson rj = new ReadJson() ;
			tabGroup tg = rj.readJSONConfigFile(configJSON, xlsGroupFile);

			return tg ;
		} catch (Exception e) {
			System.err.println("readFromJSON::Exception::" + e.getMessage()) ;
			return null;
		}
	}

	public void ReadXLSBuildGroup2(String configJSON, String xlsGroupFile) {
		try {
			tabGroup tg = readFromJSON(configJSON, xlsGroupFile);
			if (tg == null) return ;

			File hFile = InitializeXLS(xlsGroupFile);
			if (hFile == null) return ;

			boolean b = buildGroupXLS2(hFile, tg);
		} catch (Exception e) {
			System.err.println("ReadXLSBuildGroup::Exception: " + e.getMessage());
		}
	}
}