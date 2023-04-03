package gpGroupXLS;

//import gpGroupXLS.utils.fileUtils;
import gpGroupXLS.TabSummary.tabEntry;
import gpGroupXLS.xchg.ExchangeRateTable;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormat;
//import org.apache.commons.lang3.math.NumberUtils;

public class groupXLS {
	private boolean m_bHeaderCopied = false ;

	private TabSummary populateGroup(String xlsGroupFile) {
		try {
			String[][] groups = {
					{ "NovVacation.xlsx", "nov.usd", "1.0", 		"eur",	"_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",	"usd" },
					{ "NovVacation.xlsx", "nov.eur", "1.08642000", 	"eur",	"[$EUR-x-euro2] #,##0.00_);([$EUR-x-euro2] #,##0.00)",	"gbp" },
					{ "NovVacation.xlsx", "nov.gbp", "1.235255", 	"gbp",	"[$GBP-en-GB] #,##0.00",								"inr" },
					{ "NovVacation.xlsx", "nov.inr", "0.01217546",	"eur",	"[$INR] #,##0.00",										"usd" },
					{ "NovVacation.xlsx", "nov.mad", "0.097809076",	"mad",	"[$MAD] #,##0.00_);([$MAD] #,##0.00)",					"usd" }
			};
			TabSummary ts = new TabSummary() ;
			ts.setXLSFileName(xlsGroupFile) ;
			for (int i = 0; i < groups.length; i++) {
				Double dRate = Double.parseDouble(groups[i][2]);
				ts.addItem(groups[i][0], groups[i][1], dRate , groups[i][3], groups[i][4], groups[i][5]) ;
			}
			//System.out.println(ts.getXLSFileName() + "::") ; ts.dump() ;
			return ts ;
		} catch (Exception e) {
			System.err.println("Exception::" + e.getMessage()) ;
			return null ;
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

	private void buildExchangeGroup(String groupName, /*XSSFWorkbook workBookGroup, */String format) {
	}

	private void copyContents(XSSFWorkbook wbSummary, XSSFSheet sheetSummary, XSSFSheet sheetToBeGrouped, String format) {
		final int numHeaders = 2 ;
		//int row = sheetSummary.getLastRowNum() ;
		//System.out.println("sheetName|nLastRowNum:" + "\t" + sheetSummary.getSheetName() + "|" +  row);
		//row = ((row == -1) ? 0 : (++row)) ;
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
	}

	private void extractSheetData(XSSFWorkbook workBookIn, String groupName, XSSFWorkbook workBookGroup, String format) {
		XSSFSheet sheetToBeGrouped = workBookIn.getSheet(groupName) ;
		if (sheetToBeGrouped == null) return ;

		final String _SHEETNAME = "Summary" ;
		XSSFSheet sheetSummary = workBookGroup.getSheet(_SHEETNAME) ;
		if (sheetSummary == null) sheetSummary = workBookGroup.createSheet(_SHEETNAME);

		copyContents(workBookGroup, sheetSummary, sheetToBeGrouped, format) ;
	}

	/*private void extractSheetsData(XSSFWorkbook workBookIn, String groupName, XSSFWorkbook workBookGroup) {
		extractSheetData(workBookIn, groupName, workBookGroup) ;
	}*/

	private boolean buildGroupXLS2(File hXLS, TabSummary ts) {
		try {
			File fGroup = new File(hXLS.getName());
			FileInputStream fileGroup = new FileInputStream(fGroup);
			XSSFWorkbook workBookGroup = new XSSFWorkbook(fileGroup);

			XSSFWorkbook workBookIn = null;
			for (tabEntry gItem : ts.hsTabSummary) {
				File fIn = new File(gItem.fileName);
				FileInputStream fileIn = new FileInputStream(fIn);
				workBookIn = new XSSFWorkbook(fileIn);
				extractSheetData(workBookIn, gItem.groupName, workBookGroup, gItem.format) ;
				fileIn.close() ;
			}

			for (tabEntry gItem : ts.hsTabSummary) {
				buildExchangeGroup(gItem.groupName, /*workBookGroup,*/ gItem.format) ;
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


	private boolean buildGroupXLS(File hXLS, TabSummary ts) {
		try {
			for (tabEntry gItem : ts.hsTabSummary) {
				File fIn = new File(gItem.fileName);
				FileInputStream fileIn = new FileInputStream(fIn);
				XSSFWorkbook workBookIn = new XSSFWorkbook(fileIn);

				File fGroup = new File(hXLS.getName());
				FileInputStream fileGroup = new FileInputStream(fGroup);
				XSSFWorkbook workBookGroup = new XSSFWorkbook(fileGroup);

				//System.out.println("group:"+ gItem.groupName);
				extractSheetData(workBookIn, gItem.groupName, workBookGroup, gItem.format) ;

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

				fileIn.close() ;
				fileGroup.close() ;
			}
			return true ;
		} catch (FileNotFoundException nfe) {
			nfe.printStackTrace();
			return false ;
		} catch (Exception e) {
			e.printStackTrace();
			return false ;
		}
	}

	public void ReadXLSBuildGroup(String xlsGroupFile) {
		try {
			TabSummary ts = populateGroup(xlsGroupFile) ;
			File hFile = InitializeXLS(xlsGroupFile);
			if (hFile == null) return ;

			ExchangeRateTable ert = new ExchangeRateTable();
			ert.buildExchangeRateTable(ts) ;

			boolean b = buildGroupXLS(hFile, ts);

		} catch (Exception e) {
			System.out.println("Error: " + e.getMessage());
		}
	}
}