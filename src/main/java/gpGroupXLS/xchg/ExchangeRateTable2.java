package gpGroupXLS.xchg;

import gpGroupXLS.currency.currencyInfo;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;

public class ExchangeRateTable2 {
	public LinkedHashMap<String, targetCurrencies> m_targetGrid;

	private String m_SheetName = "" ;
	private ArrayList<Integer> rowRefIndex = new ArrayList<Integer>() ;
	private ArrayList<Integer> colRefIndex = new ArrayList<Integer>() ;

	public void addRates(String[] fC, String tC, String[] rates, String format) {
		if (m_targetGrid == null) m_targetGrid = new LinkedHashMap<String, targetCurrencies>();
		targetCurrencies tCur = new targetCurrencies(fC, tC, rates, format);
		m_targetGrid.put(tC, tCur);
	}

	public void addRates2(String[] fC, String tC, String format, RateDate[] rd) {
		if (m_targetGrid == null) m_targetGrid = new LinkedHashMap<String, targetCurrencies>();
		targetCurrencies tCur = new targetCurrencies(fC, tC, format, rd);
		m_targetGrid.put(tC, tCur);
	}

	public void buildRateReferenceGrid(CellReference cr) {
		m_SheetName = cr.getCellRefParts()[0] ;	// location of "from|to"
		int startRowReference = Integer.parseInt(cr.getCellRefParts()[1]);	//1;
		char startColReference = cr.getCellRefParts()[2].charAt(0);	//'A';

		int row = startRowReference ;
		char col = startColReference;

		boolean bRowsAdded = false ;
		for (Map.Entry<String, targetCurrencies> tC : m_targetGrid.entrySet()) {
			String toCurrency = tC.getKey();

			targetCurrencies fCs = tC.getValue();
			row = startRowReference ;
			col = (char) (col + 1);
			int i = col ;

			colRefIndex.add(i) ;
			if (!bRowsAdded) {
				for (exchangePair ep : fCs.m_targetRates) {
					rowRefIndex.add(++row) ;
					//colRefCurrencies.add(toCurrency) ;
				}
				bRowsAdded = true;
			}
			col = (char) (col + 1); // enhancement: date added to rate
		}
	}

	public String getRateReference(int fromCurrency, int toCurrency) {
		int rRefIndex = rowRefIndex.get(fromCurrency) ;
		int cRefIndex = colRefIndex.get(toCurrency) ;
		char c = (char) cRefIndex ;
		return m_SheetName + "!" + c + rRefIndex  ;	// "Sheet2!A1"
	}

	public void getXRates(String sheetName) {
		int r = 0 ;
		for (Map.Entry<String, targetCurrencies> tC : m_targetGrid.entrySet()) {
			String toCurrency = tC.getKey();
			int c = 0 ;
			targetCurrencies fCs = tC.getValue();
			for (exchangePair ep : fCs.m_targetRates) {
				String xRef = getRateReference(c, r) ;
				System.out.println("xRef::" + xRef) ;
				c++ ;
			}
			r++ ;
		}
	}

	public void dump() {
		for (Map.Entry<String, targetCurrencies> tC : m_targetGrid.entrySet()) {
			String toCurrency = tC.getKey();
			System.out.println(toCurrency + "::") ;
			targetCurrencies fCs = tC.getValue();
			fCs.dump() ;
		}
	}

	public class targetCurrencies {
		public ArrayList<exchangePair> m_targetRates;
		public currencyInfo m_CurrencyInfo;

		public targetCurrencies(String[] fC, String tC, String[] rates, String format) {
			if (fC.length != rates.length) return;
			m_targetRates = new ArrayList<exchangePair>() ;
			for (int i = 0; i < rates.length; i++) {
				Double r = Double.parseDouble(rates[i]);
				String f = fC[i];
				exchangePair ep = new exchangePair(f, tC, r);
				boolean b = m_targetRates.add(ep);
			}
			m_CurrencyInfo = new currencyInfo(tC, format);
		}

		public targetCurrencies(String[] fC, String tC, String format, RateDate[] rd) {
			if (fC.length != rd.length) return;
			m_targetRates = new ArrayList<exchangePair>() ;
			for (int i = 0; i < rd.length; i++) {
				Double r = rd[i].rate;
				String d = rd[i].date;
				String f = fC[i];
				exchangePair ep = new exchangePair(f, tC, r, d);
				boolean b = m_targetRates.add(ep);
			}
			m_CurrencyInfo = new currencyInfo(tC, format);
		}

		public void dump() {
			for (exchangePair ep : m_targetRates) {
				System.out.println(ep);
			}
		}
	}

	public class exchangePair {
		public String	fromCurrency ;	// = groupTabs.tabEntry.currency + ":" + targetGrid.key
		public String	toCurrency ;	// = groupTabs.tabEntry.currency + ":" + targetGrid.key
		public Double	rate;
		public RateDate	rd ;

		public exchangePair(String fC, String tC, Double r) {
			fromCurrency = fC;
			toCurrency = tC ;
			rate = r ;
			rd = null;
		}

		public exchangePair(String fC, String tC, Double r, String d) {
			fromCurrency = fC;
			toCurrency = tC ;
			rd = new RateDate(r, d) ;
			rate = 0.0D;
		}

		@Override public String toString() {
			final String _SEP = "|" ;
			return "exchangePair [" + this.fromCurrency + _SEP + this.toCurrency + _SEP + this.rate + _SEP + this.rd.toString() + "]";
		}
	}

	public class RateDate {
		public Double	rate;
		public String 	date;
	
		public RateDate(Double r, String d) {
			rate = r ;
			date = d ;
		}		

		@Override public String toString() {
			final String _SEP = "|" ;
			return "RateDate [" + this.rate + _SEP + this.date + "]";
		}
	}
}
