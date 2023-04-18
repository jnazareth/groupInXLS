package gpGroupXLS.group;

import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.xchg.ExchangeRateTable2;

public class tabGroup {
	public TabSummary2	m_groupTabs = new TabSummary2();
    public ExchangeRateTable2 m_xTable = new ExchangeRateTable2(); 

	public tabGroup(TabSummary2 ts2, ExchangeRateTable2 ert2) {
		if (ts2 != null) m_groupTabs = ts2;
		if (ert2 != null) m_xTable = ert2;
	}
}
