package gpGroupXLS.json;

import gpGroupXLS.utils.fileUtils;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.xchg.ExchangeRateTable2;
import gpGroupXLS.group.tabGroup;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.util.Map;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.ArrayList;

public class ReadJson {
    public tabGroup readJSONConfigFile(String configFile, String xlsGroupFile) {
		tabGroup tg = null ;

        JSONParser jsonParser = new JSONParser();
        try {
			FileReader reader = fileUtils.getFileReader(configFile) ;

            //Read JSON file
            Object oParser = jsonParser.parse(reader);
            org.json.simple.JSONObject jo = (org.json.simple.JSONObject) oParser;

            //String fileName  = (String) jo.get(JSONKeys.keyFileName);
            /*Iterator<String> keys = jo.keySet().iterator();
            while (keys.hasNext()) {
                System.out.println("value: " + keys.next());
            }*/

			TabSummary2 ts = new TabSummary2() ;
			ts.setXLSFileName(xlsGroupFile) ;

            String cd  = (String) jo.get(JSONKeys.keySumColumns);
			ts.setCoords(cd) ;

            JSONArray joGroupTabs ;
            joGroupTabs = (JSONArray)jo.get(JSONKeys.keyGrouptabs);
			ArrayList<String> fCurrencies = new ArrayList<String>(joGroupTabs.size()) ;
            for (int i = 0; i < joGroupTabs.size(); i++) {
                JSONObject item = (JSONObject)joGroupTabs.get(i);
                String fName = (String)item.get(JSONKeys.keyFileName);
                String gName = (String)item.get(JSONKeys.keyGroupName);
                String fCur = (String)item.get(JSONKeys.keyCurrency);
                String sFormat = (String)item.get(JSONKeys.keyFormat);
                //System.out.println("fName:" + fName + "\t\tgName:" + gName + "\t\tsCurrency:" + fCur + "\t\tsFormat:" + sFormat + "\t\tsSumColumns:" + sSumColumns);

				fCurrencies.add(fCur) ;
				//ts.addItem(fName, gName, fCur, sFormat, sSumColumns) ;
				ts.addItem(fName, gName, fCur, sFormat) ;
            }
            //ts.dump() ;

			ExchangeRateTable2 xrt2 = new ExchangeRateTable2();
            JSONArray joTargetCurrrencies ;
            joTargetCurrrencies = (JSONArray)jo.get(JSONKeys.keyTargetCurrrencies);
            for (int i = 0; i < joTargetCurrrencies.size(); i++) {
                JSONObject item = (JSONObject)joTargetCurrrencies.get(i);
                String toCurrency = (String)item.get(JSONKeys.keyCurrency);
                String sFormat = (String)item.get(JSONKeys.keyFormat);
                //System.out.println("toCurrency:" + toCurrency + "\t\tsFormat:" + sFormat);

	            JSONArray joRates ;
	            joRates = (JSONArray)item.get(JSONKeys.keyRates);
				ArrayList<String> arrayRates = new ArrayList<String>(joRates.size()) ;
	            for (int j = 0; j < joRates.size(); j++) {
					String rate = (String)joRates.get(j);
	                arrayRates.add(rate) ;
				}

				String[] fromCurrency = new String[fCurrencies.size()];
				fromCurrency = fCurrencies.toArray(fromCurrency) ;

				String[] arates = new String[arrayRates.size()];
				arates = arrayRates.toArray(arates) ;

				xrt2.addRates(fromCurrency, toCurrency, arates, sFormat);            }
				//xrt2.dump() ;

			reader.close();
			tg = new tabGroup(ts, xrt2);
			return tg;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        } catch (ParseException e) {
            e.printStackTrace();
            return null;
        }
    }
}