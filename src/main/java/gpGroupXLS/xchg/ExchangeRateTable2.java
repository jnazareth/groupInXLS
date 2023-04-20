package gpGroupXLS.xchg;

import gpGroupXLS.currency.currencyInfo;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExchangeRateTable2 {
    public LinkedHashMap<String, targetCurrencies> m_targetGrid;

    public void addRates(String[] fC, String tC, String[] rates, String format) {
        if (m_targetGrid == null) m_targetGrid = new LinkedHashMap<String, targetCurrencies>();
        targetCurrencies tCur = new targetCurrencies(fC, tC, rates, format);
        m_targetGrid.put(tC, tCur);
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

		public void dump() {
			for (exchangePair ep : m_targetRates) {
				ep.dump() ;
			}
		}

    }

    public class exchangePair {
		public String	fromCurrency ;	// = groupTabs.tabEntry.currency + ":" + targetGrid.key
		public String	toCurrency ;	// = groupTabs.tabEntry.currency + ":" + targetGrid.key
		public Double	rate;

        public exchangePair(String fC, String tC, Double r) {
            fromCurrency = fC;
            toCurrency = tC ;
            rate = r ;
        }

        public void dump() {
			final String _SEP = "|" ;
			System.out.println(fromCurrency + _SEP + toCurrency + _SEP + rate) ;
        }
    }
}
