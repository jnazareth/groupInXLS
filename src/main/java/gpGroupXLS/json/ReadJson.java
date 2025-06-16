package gpGroupXLS.json;

import gpGroupXLS.utils.fileUtils;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.xchg.ExchangeRateTable;
import gpGroupXLS.xchg.ExchangeRateTable.RateDate;
import gpGroupXLS.group.TabGroup;

import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class ReadJson {

    public TabGroup readJSONConfigFile(String configFile, String xlsGroupFile) {
        JSONParser jsonParser = new JSONParser();

        try (FileReader reader = fileUtils.getFileReader(configFile)) {
            JSONObject jsonObject = (JSONObject) jsonParser.parse(reader);

			List<String> fCurrencies = new ArrayList<>();
            TabSummary2 tabSummary = parseTabSummary(jsonObject, xlsGroupFile, fCurrencies);
            ExchangeRateTable exchangeRateTable = parseExchangeRateTable(jsonObject, tabSummary, fCurrencies);

            //System.out.println("readJSONConfigFile::" + exchangeRateTable);

            return new TabGroup(tabSummary, exchangeRateTable);
        } catch (IOException | ParseException e) {
            e.printStackTrace();
            return null;
        }
    }

    private TabSummary2 parseTabSummary(JSONObject jsonObject, String xlsGroupFile, List<String> fCurrencies) {
        TabSummary2 tabSummary = new TabSummary2();
        tabSummary.setXLSFileName(xlsGroupFile);

        String sumColumns = (String) jsonObject.get(JSONKeys.SUM_COLUMNS);
        tabSummary.setCoords(sumColumns);

        long numPersons = (long) jsonObject.get(JSONKeys.NUM_PERSONS);
        tabSummary.setNumPersons((int) numPersons);

        JSONArray groupTabs = (JSONArray) jsonObject.get(JSONKeys.GROUP_TABS);

        for (Object obj : groupTabs) {
            JSONObject item = (JSONObject) obj;
            String fileName = (String) item.get(JSONKeys.FILE_NAME);
            String groupName = (String) item.get(JSONKeys.GROUP_NAME);
            String currency = (String) item.get(JSONKeys.CURRENCY);
            String format = (String) item.get(JSONKeys.FORMAT);

            fCurrencies.add(currency);
            tabSummary.addItem(fileName, groupName, currency, format);
        }
        return tabSummary;
    }

    private ExchangeRateTable parseExchangeRateTable(JSONObject jsonObject, TabSummary2 tabSummary, List<String> fromCurrencies) {
        ExchangeRateTable exchangeRateTable = new ExchangeRateTable();
        JSONArray targetCurrencies = (JSONArray) jsonObject.get(JSONKeys.TARGET_CURRENCIES);

        for (Object obj : targetCurrencies) {
            JSONObject item = (JSONObject) obj;
            String toCurrency = (String) item.get(JSONKeys.CURRENCY);
            String format = (String) item.get(JSONKeys.FORMAT);

            JSONArray ratesArray = (JSONArray) item.get(JSONKeys.RATES);
            RateDate[] rateDates = new RateDate[ratesArray.size()];

            for (int i = 0; i < ratesArray.size(); i++) {
                JSONObject rateItem = (JSONObject) ratesArray.get(i);
                double rate = Double.parseDouble((String) rateItem.get(JSONKeys.RATE));
                String date = (String) rateItem.get(JSONKeys.DATE);
                rateDates[i] = exchangeRateTable.new RateDate(rate, date);
            }


			String[] fromCurrency = new String[fromCurrencies.size()];
			fromCurrency = fromCurrencies.toArray(fromCurrency) ;
            exchangeRateTable.addRates(fromCurrency, toCurrency, format, rateDates);
        }

        //System.out.println("parseExchangeRateTable::" + exchangeRateTable);
        return exchangeRateTable;
    }

    @Override
    public String toString() {
        return "ReadJson{}";
    }
}