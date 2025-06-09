package gpGroupXLS.group;

import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.xchg.ExchangeRateTable;

public class tabGroup {
	public TabSummary2	m_tabSummary = new TabSummary2();
    public ExchangeRateTable m_xTable = new ExchangeRateTable();

	public tabGroup(TabSummary2 ts, ExchangeRateTable ert) {
		if (ts != null) m_tabSummary = ts;
		if (ert != null) m_xTable = ert;
	}
}