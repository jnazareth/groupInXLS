package gpGroupXLS.group;

import gpGroupXLS.tabs.TabSummary;
import gpGroupXLS.xchg.ExchangeRateTable;

public class tabGroup {
	public TabSummary	m_tabSummary = new TabSummary();
    public ExchangeRateTable m_xTable = new ExchangeRateTable();

	public tabGroup(TabSummary ts, ExchangeRateTable ert) {
		if (ts != null) m_tabSummary = ts;
		if (ert != null) m_xTable = ert;
	}
}
