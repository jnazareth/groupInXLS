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


public class groupXLS {
	private final String _SHEETNAME = "Summary" ;
	private boolean m_bHeaderCopied = false ;

	private TabSummary2 populateGroup2(String xlsGroupFile) {
		try {
			String[][] groups = {
					{ "NovVacation.xlsx", "nov.usd", "usd",	"_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)", "[C9:C20]"},
					{ "NovVacation.xlsx", "nov.eur", "eur",	"[$EUR-x-euro2] #,##0.00_);([$EUR-x-euro2] #,##0.00)", "[C9:C20]"},
					{ "NovVacation.xlsx", "nov.gbp", "gbp",	"[$GBP-en-GB] #,##0.00", "[C9:C20]"},
					{ "NovVacation.xlsx", "nov.inr", "inr",	"[$INR] #,##0.00", "[C9:C20]"},
					{ "NovVacation.xlsx", "nov.mad", "mad",	"[$MAD] #,##0.00_);([$MAD] #,##0.00)", "[C9:C20]"}
			};
			TabSummary2 ts = new TabSummary2() ;
			ts.setXLSFileName(xlsGroupFile) ;
			for (int i = 0; i < groups.length; i++) {
				ts.addItem(groups[i][0], groups[i][1], groups[i][2], groups[i][3], groups[i][4]) ;
			}
			return ts ;
		} catch (Exception e) {
			System.err.println("populateGroup2::Exception::" + e.getMessage()) ;
			return null ;
		}
	}

	private ExchangeRateTable2 populateRates2(String xlsGroupFile, TabSummary2 ts2) {
		try {
			String[][] rates = {
				/* usd */ {"1", "1.08642", "1.235255", "0.01217546", "0.097809076"},
				/* eur */ {"0.920454336", "1", "1.136995821", "0.011206955", "0.090028788"},
				/* gbp */ {"0.809549445", "0.879510708", "1", "0.009856637", "0.079181283"}
			};
			ArrayList<String> cCrossCurrencies = new ArrayList<String>();
			// target currency index *must* match exchange rates array above.
			cCrossCurrencies.add(0, "usd");
			cCrossCurrencies.add(1, "eur");
			cCrossCurrencies.add(2, "gbp");

			ArrayList<String> cCurrencyFormats = new ArrayList<String>();
			// target currency index *must* match exchange rates array above.
			cCurrencyFormats.add(0, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)");
			cCurrencyFormats.add(1, "[$EUR-x-euro2] #,##0.00_);([$EUR-x-euro2] #,##0.00)");
			cCurrencyFormats.add(2, "[$GBP-en-GB] #,##0.00");

			ExchangeRateTable2 xrt2 = new ExchangeRateTable2();
			for (String tC : cCrossCurrencies) {
				int targetCurrencyIndex = cCrossCurrencies.indexOf(tC);
				String[] fC = new String[rates[targetCurrencyIndex].length];
				for (int k = 0; k < rates[targetCurrencyIndex].length; k++) {
					tabGroupBase tb = ts2.m_groupTabs.get(k);
					fC[k] = tb.te.currency;
				}
				xrt2.addRates(fC, tC, rates[targetCurrencyIndex], cCurrencyFormats.get(targetCurrencyIndex));
			}
			return xrt2;
		} catch (Exception e) {
			System.err.println("populateRates2::Exception::" + e.getMessage()) ;
			return null;
		}
	}

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
		//if (newRow == null) return ;

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
		//if (newRow == null) return ;

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
		//System.out.println("newRow.getRowNum():" + newRow.getRowNum());

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

	private void readTables(XSSFSheet sheet){
		System.out.println("looking for tables in " + sheet.getSheetName());
		List<XSSFTable> tables = sheet.getTables();
		for (XSSFTable t : tables) {
			System.out.println(t.getDisplayName());
			System.out.println(t.getName());
			//System.out.println(t.getNumerOfMappedColumns());

			int startRow = t.getStartCellReference().getRow();
			int endRow = t.getEndCellReference().getRow();
			System.out.println("startRow = " + startRow);
			System.out.println("endRow = " + endRow);

			int startColumn = t.getStartCellReference().getCol();
			int endColumn = t.getEndCellReference().getCol();

			System.out.println("startColumn = " + startColumn);
			System.out.println("endColumn = " + endColumn);
		}
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

	private int copyContents2(XSSFWorkbook wbSummary, XSSFSheet sheetSummary, XSSFSheet sheetToBeGrouped, String format) {
		final int numHeaders = 2 ;
		int row = sheetSummary.getLastRowNum() + 1 ;
		if (row >= numHeaders) m_bHeaderCopied = true;

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		int rowToBeGrouped = 0 ;
		Iterator<Row> aRowIterator = sheetToBeGrouped.rowIterator();
		while (aRowIterator.hasNext()) {
		//Row rowitr = sheetToBeGrouped.getRow(lRow) ;
		//if (rowitr != null)
			Row rowitr = (Row) aRowIterator.next();

			if ((m_bHeaderCopied) && (rowToBeGrouped < numHeaders)) {
				; // do nothing
			} else {
				Row currentRow = sheetSummary.createRow(row);

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
			rowToBeGrouped++ ;
		}
		return (row-1) ;
	}

	/*private int copyContents(XSSFWorkbook wbSummary, XSSFSheet sheetSummary, XSSFSheet sheetToBeGrouped, String format) {
		final int numHeaders = 2 ;
		int row = sheetSummary.getLastRowNum() + 1 ;
		if (row >= numHeaders) m_bHeaderCopied = true;

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		int rowToBeGrouped = 0 ;
		Iterator<Row> aRowIterator = sheetToBeGrouped.rowIterator();
		while (aRowIterator.hasNext()) {
			Row rowitr = (Row) aRowIterator.next();

			if ((m_bHeaderCopied) && (rowToBeGrouped < numHeaders)) {
				; // do nothing
			} else {
				Row currentRow = sheetSummary.createRow(row);

				Iterator<Cell> aCellIterator = rowitr.cellIterator();
				int col = 0 ;
				while(aCellIterator.hasNext()) {
					Cell celldata = (Cell) aCellIterator.next();
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
					col++ ;
				}
			}
			row = sheetSummary.getLastRowNum() + 1 ;
			rowToBeGrouped++ ;
		}
		return (row-1) ;
	}*/

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

	/*private void extractSheetsData(XSSFWorkbook workBookIn, String groupName, XSSFWorkbook workBookGroup) {
		extractSheetData(workBookIn, groupName, workBookGroup) ;
	}*/

	private boolean buildGroupXLS2(File hXLS, tabGroup tg) {
		try {
			File fGroup = new File(hXLS.getName());
			FileInputStream fileGroup = new FileInputStream(fGroup);
			XSSFWorkbook workBookGroup = new XSSFWorkbook(fileGroup);

			// build base grid
			boolean bFirst = true ;
			XSSFWorkbook workBookIn = null;
			TabSummary2 ts2 = tg.m_groupTabs ;
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

			ExchangeRateTable2 ert2 = tg.m_xTable;
			HashMap<String, targetCurrencies> tGrid = ert2.m_targetGrid;
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
					tabGroupBase tgbase = tg.m_groupTabs.m_groupTabs.get(i);
					//String fromCurrency = tR.get(i).fromCurrency;
					Double rate = tR.get(i).rate;
					//_Coordinates cd = tgbase.te.coords;
					int row = tgbase.rowNumber ;
					if (row != -1) {
						int rA = buildExchangeGroup2(workBookGroup, toCurrency, row, rate, cFormat) ;
						if (rA != -1) rowsAdded.add(rA) ;
					}
				}

				// add sum
				_Coordinates cd = null;
				for (int i = 0; i < tR.size(); i++) {
					tabGroupBase tgbase = tg.m_groupTabs.m_groupTabs.get(i);
					cd = tgbase.te.coords;
				}
				buildExchangeGroupSum(workBookGroup, toCurrency, rowsAdded, cd, cFormat) ;
			}

			try {
				workBookIn.close();
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
			tabGroup tg = null ;

			ReadJson rj = new ReadJson() ;
			tg = rj.readJSONConfigFile(configJSON, xlsGroupFile);

			return tg ;
		} catch (Exception e) {
			System.err.println("readFromJSON::Exception::" + e.getMessage()) ;
			return null;
		}
	}

	public void ReadXLSBuildGroup(String xlsGroupFile) {
		try {
			TabSummary2 ts2 = populateGroup2(xlsGroupFile) ;
			ExchangeRateTable2 ert2 = populateRates2(xlsGroupFile, ts2) ;
			tabGroup tg = new tabGroup(ts2, ert2);

			File hFile = InitializeXLS(xlsGroupFile);
			if (hFile == null) return ;

			boolean b = buildGroupXLS2(hFile, tg);
		} catch (Exception e) {
			System.err.println("ReadXLSBuildGroup::Exception: " + e.getMessage());
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