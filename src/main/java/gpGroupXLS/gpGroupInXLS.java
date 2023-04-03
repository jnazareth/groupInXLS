package gpGroupXLS;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

public class gpGroupInXLS {
	public static void main(String[] args) {
		try {
			gpGroupInXLS app = new gpGroupInXLS();

			int parmNo;
			String xlsGroupFile = "" ;

			//command line optional parameters
			if (args.length == 0 || args.length > 2) {
				//app.showUsage() ;
				return;
			}

			for (parmNo = 0; parmNo < args.length; parmNo++) {
				if (args[parmNo].substring(0, 2) == "-h") {
					//app.showUsage() ;
					return;
				}
				else {
					if (parmNo == 0) xlsGroupFile = args[parmNo] ;
				}
			}

			groupXLS grpXLS = new groupXLS() ;
			grpXLS.ReadXLSBuildGroup(xlsGroupFile) ;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
