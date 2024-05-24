package gpGroupXLS.group;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.format.Export;
import gpGroupXLS.format.Export.RowLayout;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.tabs.TabSummary2.tabEntry2;
import gpGroupXLS.tabs.TabSummary2.tabGroupBase;

public class BaseGrid {
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

	private int copyContents3(XSSFWorkbook wbSummary, XSSFSheet sheetSummary, String groupName, XSSFSheet sheetToBeGrouped, String format, LinkedHashSet<Integer> rowsToCopy) {
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
		rowsToCopy2 = (LinkedHashSet<Integer>) rowsToCopy.clone();
		if (sourceRow != -1) rowsToCopy2.add(sourceRow) ;
		Integer[] sourceRows = new Integer [rowsToCopy2.size()];
		sourceRows = rowsToCopy2.toArray(sourceRows) ;

		//createStyle
		XSSFCellStyle cellStyle = wbSummary.createCellStyle();
		DataFormat dFormat = wbSummary.createDataFormat();
		cellStyle.setDataFormat(dFormat.getFormat(format));

		// _newFormatting 4/22
		String gName = "";
		Integer rNum = -1 ;
		for (int r = 0; r < sourceRows.length; r++) {
			// _newFormatting 4/22
			if (r == (sourceRows.length-1)) gName = groupName;
			int colInsertPosition = 0;

			Row rowitr = sheetToBeGrouped.getRow(sourceRows[r]) ;
			if (rowitr != null) {
				Row currentRow = sheetSummary.createRow(row);
				rowAdded = currentRow.getRowNum();

				int lastColumn = rowitr.getLastCellNum() ;
				// _newFormatting 4/22
				for (int col = (XLSProperties._numberToSkip - XLSProperties._groupNameColOffset); col < lastColumn; col++) {
					// _newFormatting 4/22
					// col to add group name
					if (col == (XLSProperties._numberToSkip - XLSProperties._groupNameColOffset)) {
						Cell cT = currentRow.createCell(colInsertPosition++);
						cT.setCellValue(gName);
						continue ;
					}

					Cell celldata = rowitr.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					if (celldata == null) {
						// The spreadsheet is empty in this cell
					} else {
						Cell cellTarget = null;
						switch(celldata.getCellType()) {
							case STRING:
								String cV = celldata.getStringCellValue() ;
								cellTarget = currentRow.createCell(colInsertPosition++);
								cellTarget.setCellValue(cV);
								break;
							case NUMERIC:
								Double d = celldata.getNumericCellValue() ;
								cellTarget = currentRow.createCell(colInsertPosition++);
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

		XSSFSheet sheetSummary = workBookGroup.getSheet(XLSProperties._SummarySheetName) ;
		if (sheetSummary == null) sheetSummary = workBookGroup.createSheet(XLSProperties._SummarySheetName);

		LinkedHashSet<Integer> rowsToCopy = new LinkedHashSet<Integer>();
		if (bFirst) {
			rowsToCopy.add(0) ;		// header# 1 (row 0)
			rowsToCopy.add(1) ;		// header# 2 (row 1)
		}
		int rInserted = copyContents3(workBookGroup, sheetSummary, groupName, sheetToBeGrouped, format, rowsToCopy) ;
		return rInserted;
	}

	public void buildBaseGrid(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			XSSFWorkbook workBookIn = null;
			TabSummary2 ts2 = tg.m_tabSummary ;

			boolean bFirst = true ;
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
}
