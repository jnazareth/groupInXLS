package gpGroupXLS;

import gpGroupXLS.group.groupXLS;


public class gpGroupInXLS {
	public static void main(String[] args) {
		try {
			gpGroupInXLS app = new gpGroupInXLS();

			int parmNo;
			String xlsGroupFile = "", configJSON = "" ;

			//command line optional parameters
			if (args.length == 0 || args.length > 3) {
				//app.showUsage() ;
				return;
			}

			for (parmNo = 0; parmNo < args.length; parmNo++) {
				if (args[parmNo].substring(0, 2) == "-h") {
					//app.showUsage() ;
					return;
				}
				else {
					if (parmNo == 0) configJSON = args[parmNo] ;
					if (parmNo == 1) xlsGroupFile = args[parmNo] ;
				}
			}

			groupXLS grpXLS = new groupXLS() ;
			grpXLS.ReadXLSBuildGroup2(configJSON, xlsGroupFile) ;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
