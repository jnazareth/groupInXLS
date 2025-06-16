package gpGroupXLS.currency;

import java.util.HashMap;
import java.util.Map;

public class CurrencyInfo {
    private final Map<String, String> currencyFormats;

    public CurrencyInfo(String currency, String format) {
        currencyFormats = new HashMap<>();
        addCurrencyFormat(currency, format);
    }

    public void addCurrencyFormat(String currency, String format) {
        currencyFormats.put(currency, format);
    }

    public String getCurrencyFormat(String currency) {
        return currencyFormats.get(currency);
    }
}