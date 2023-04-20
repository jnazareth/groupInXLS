package gpGroupXLS.currency;

import java.util.HashMap;

public class currencyInfo {
    private HashMap<String, String> m_CI = new HashMap<String, String>();

    public currencyInfo(String c, String f) {
        m_CI.put(c, f);
    }

    public String getCurrencyFormat(String currency) {
        return m_CI.get(currency);
    }
}
