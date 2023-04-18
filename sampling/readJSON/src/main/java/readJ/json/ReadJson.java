package readJ.json;

import readJ.utils.fileUtils;

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

public class ReadJson {
    public void readJSONConfigFile(String configFile) {
        //groupCsvJsonMapping gMapping = new groupCsvJsonMapping() ;

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

            JSONArray joGroupTabs ;
            joGroupTabs = (JSONArray)jo.get(JSONKeys.keyGrouptabs);
            for (int i = 0; i < joGroupTabs.size(); i++) {
                //System.out.println(joGroupTabs.get(i));
                JSONObject item = (JSONObject)joGroupTabs.get(i);
                String fName = (String)item.get(JSONKeys.keyFileName);
                String gName = (String)item.get(JSONKeys.keyGroupName);
                String sCurrency = (String)item.get(JSONKeys.keyCurrency);
                String sFormat = (String)item.get(JSONKeys.keyFormat);
                String sSumColumns = (String)item.get(JSONKeys.keySumColumns);
                System.out.println("fName:" + fName + "\t\tgName:" + gName + "\t\tsCurrency:" + sCurrency + "\t\tsFormat:" + sFormat + "\t\tsSumColumns:" + sSumColumns);

                //csvFileJSON cj = new csvFileJSON();
                //_SheetProperties sp = new _SheetProperties() ;
                //gMapping.addItem(gName, sFile, sJSON, cj, sp);
            }

            JSONArray joTargetCurrrencies ;
            joTargetCurrrencies = (JSONArray)jo.get(JSONKeys.keyTargetCurrrencies);
            for (int i = 0; i < joTargetCurrrencies.size(); i++) {
                //System.out.println(joTargetCurrrencies.get(i));
                JSONObject item = (JSONObject)joTargetCurrrencies.get(i);
                String sCurrency = (String)item.get(JSONKeys.keyCurrency);
                String sFormat = (String)item.get(JSONKeys.keyFormat);
                System.out.println("sCurrency:" + sCurrency + "\t\tsFormat:" + sFormat);

	            JSONArray joRates ;
	            joRates = (JSONArray)item.get(JSONKeys.keyRates);
	            for (int j = 0; j < joRates.size(); j++) {
					String rate = (String)joRates.get(j);
	                System.out.println("sRate:" + rate);
				}
            }

            //gMapping.dumpCollection();
			reader.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }
}
