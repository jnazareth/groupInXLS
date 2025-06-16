package gpGroupXLS.group;

import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.xchg.ExchangeRateTable;

public class TabGroup {
    private final TabSummary2 tabSummary;
    private final ExchangeRateTable exchangeRateTable;

    public TabGroup(TabSummary2 tabSummary, ExchangeRateTable exchangeRateTable) {
        this.tabSummary = tabSummary != null ? tabSummary : new TabSummary2();
        this.exchangeRateTable = exchangeRateTable != null ? exchangeRateTable : new ExchangeRateTable();
    }

    public TabSummary2 getTabSummary() {
        return tabSummary;
    }

    public ExchangeRateTable getExchangeRateTable() {
        return exchangeRateTable;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("TabGroup{");
        sb.append("tabSummary=").append(tabSummary);
        sb.append(", exchangeRateTable=").append(exchangeRateTable);
        sb.append('}');
        return sb.toString();
    }
}