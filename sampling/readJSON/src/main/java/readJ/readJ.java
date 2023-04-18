package readJ;

import readJ.json.ReadJson;

public class readJ {
	public static void main(String[] args) {
		System.out.println("hello readJ") ;

		ReadJson rj = new ReadJson() ;

		String configFile = "config1.json";
		rj.readJSONConfigFile(configFile);
	}
}
