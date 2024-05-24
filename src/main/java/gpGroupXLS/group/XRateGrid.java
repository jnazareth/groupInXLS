package gpGroupXLS.group;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.xchg.ExchangeRateTable2;
import gpGroupXLS.xchg.ExchangeRateTable2.exchangePair;
import gpGroupXLS.xchg.ExchangeRateTable2.targetCurrencies;

public class XRateGrid {
	private CellReference addXRSheet(XSSFWorkbook workBookGroup, String[][] arrayRates) {
		try {
			XSSFSheet sheetXRate = workBookGroup.getSheet(XLSProperties._XRatesSheetName) ;
			if (sheetXRate == null) sheetXRate = workBookGroup.createSheet(XLSProperties._XRatesSheetName);

			for (int i = 0; i < arrayRates.length ; i++) {
				Row newRow = sheetXRate.createRow(i);
				for (int j = 0; j < arrayRates[i].length; j++) {
					Cell cellTarget = newRow.createCell(j);
					cellTarget.setCellValue(arrayRates[i][j]);
				}
			}
			// start of table = cell "from|to"
			CellReference a1 = new CellReference(sheetXRate.getSheetName() + "!A1");
			return a1;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null ;
	}

	private void dumprates(String[][] a){
		for (int i = 0; i < a.length ; i++) {
			for (int j = 0; j < a[i].length; j++) {
				System.out.println("[" + i + "][" + j + "]:" + a[i][j]) ;
			}
		}
	}

	private String[][] transpose(String[][] array) {
		if (array == null || array.length == 0) return array;
		int width = array.length;
		int height = array[0].length;
		String[][] arrayTransposed = new String[height][width];
		for (int x = 0; x < width; x++) {
			for (int y = 0; y < height; y++) {
				arrayTransposed[y][x] = array[x][y];
			}
		}
		return arrayTransposed;
	}

	private void dumpXRates(String [][] xr){
		for (int m = 0; m < xr.length; m++) {
			for (int n = 0; n < xr[m].length; n++) {
				System.out.print("xr [" + m + "][" + n + "]:" + xr[m][n] + "\t") ;
			}
			System.out.println("") ;
		}
	}

	private void populateXRates(String[][] xg, ArrayList<String> col, final int idx) {
		if (xg.length != col.size()) return;		// size mismatch
		try{
			String[] values = new String[col.size()];
			values = col.toArray(values);
			for (int i = 0; i < xg.length; i++) {
				xg[i][idx] = values[i] ;
			}
			col.clear();
		} catch (ArrayIndexOutOfBoundsException oobe){
			System.err.println("populateXRates::Exception::" + oobe.getMessage()) ;
		}
	}

	private String[][] buildratesArrayGrid2(tabGroup tg) {
		try {
			final String _sFromTo = "from (C) | to (R)" ;
			final String _sDate = "date" ;

			ExchangeRateTable2 ert2 = tg.m_xTable;
			LinkedHashMap<String, targetCurrencies> tGrid = ert2.m_targetGrid;
			int m = tGrid.entrySet().size();
			int n = 0;
			for (Map.Entry<String, targetCurrencies> tC : tGrid.entrySet()) {
				//String toCurrency = tC.getKey();
				targetCurrencies fCs = tC.getValue();
				n = fCs.m_targetRates.size();
			}
			final int iHeader = 1 ;
			int c = (m * 2) + iHeader; 	// each entry: rate+date + header
			int r = n + iHeader;		// header
			String[][] xratesGrid = new String[r][c];

			ArrayList<String> aCol = new ArrayList<String>(c);
			ArrayList<String> aColR = new ArrayList<String>(c);
			ArrayList<String> aColD = new ArrayList<String>(c);

			ert2 = tg.m_xTable;
			tGrid = ert2.m_targetGrid;

			boolean bColHeader = true, bCurHeader = true ;
			int j = 0;
			for (Map.Entry<String, targetCurrencies> tC : tGrid.entrySet()) {
				targetCurrencies fCs = tC.getValue();
				for (exchangePair ep : fCs.m_targetRates) {
					if (j == 0) {
						if (bColHeader) {
							aCol.add(_sFromTo) ;
							bColHeader = false ;
						}
						aCol.add(ep.fromCurrency) ;
					}
					if (bCurHeader) {
						aColR.add(ep.toCurrency) ;	aColD.add(_sDate) ;
						bCurHeader = false ;
					}
					aColR.add(String.valueOf(ep.rd.rate)) ;	aColD.add(ep.rd.date) ;
				}
				bCurHeader = true;
				if (j == 0) {
					populateXRates(xratesGrid, aCol, j);
				}
				populateXRates(xratesGrid, aColR, j+1);
				populateXRates(xratesGrid, aColD, j+2);

				j += 2;
			}

			//System.out.println("after:") ;
			//dumpXRates(xratesGrid);

			return xratesGrid;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	private String[][] buildratesArrayGrid(tabGroup tg) {
		try {
			final String _sFromTo = "from (C) | to (R)" ;

			ExchangeRateTable2 ert2 = tg.m_xTable;
			LinkedHashMap<String, targetCurrencies> tGrid = ert2.m_targetGrid;
			int m = tGrid.entrySet().size();
			int n = 0;
			for (Map.Entry<String, targetCurrencies> tC : tGrid.entrySet()) {
				String toCurrency = tC.getKey();
				targetCurrencies fCs = tC.getValue();
				n = fCs.m_targetRates.size();
			}
			String[][] xratesGrid = new String[m+1][n+1];
			ArrayList<String> aRow = new ArrayList<String>(n+1);
			ArrayList<String> aCol = new ArrayList<String>(m+1);

			ert2 = tg.m_xTable;
			tGrid = ert2.m_targetGrid;
			int r = 0 ;
			boolean bRowHeader = true ;
			for (Map.Entry<String, targetCurrencies> tC : tGrid.entrySet()) {
				//String toCurrency = tC.getKey();
				targetCurrencies fCs = tC.getValue();

				aRow.clear();
				//aRow.add(toCurrency);
				boolean bColHeader = true;
				aCol.clear();
				if (bColHeader) aCol.add(_sFromTo);
				for (exchangePair ep : fCs.m_targetRates) {
					if (bColHeader) {
						aRow.add(ep.toCurrency);
						bColHeader = false;
					}
					aRow.add(String.valueOf(ep.rate));
					if (bRowHeader) {
						aCol.add(ep.fromCurrency);
					}
				}
				if (bRowHeader) {
					String[] header = new String[aCol.size()];
					xratesGrid[r] = aCol.toArray(header);
					bRowHeader = false;
					r++ ;
				}
				String[] arates = new String[aRow.size()];
				xratesGrid[r] = aRow.toArray(arates);
				r++ ;
			}
			//System.out.println("before:") ;
			//dumpXRates(xratesGrid);

			xratesGrid = transpose(xratesGrid);

			//System.out.println("after:") ;
			//dumpXRates(xratesGrid);
			return xratesGrid;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	private CellReference addXRateSheet2(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			String[][] arrayRates = buildratesArrayGrid2(tg);
			if (arrayRates == null) return null;

			CellReference cr = addXRSheet(workBookGroup, arrayRates);
			if (cr == null) return null;

			String[] crParts = cr.getCellRefParts() ;
			if ((crParts == null) || (crParts.length == 0)) return null ;

			return cr ;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null ;
	}

	public void buildXRateGrid(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			CellReference cr = addXRateSheet2(workBookGroup, tg) ;
			if (cr == null) return ;

			ExchangeRateTable2 ert2 = tg.m_xTable;
			ert2.buildRateReferenceGrid(cr);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}