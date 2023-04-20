package gpGroupXLS.utils;

import java.io.File;
import java.io.FileInputStream ;
import java.io.FileReader;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.commons.io.FileUtils ;

public class fileUtils {
	public static File getFile(String fileName)
	throws FileNotFoundException
	{
		File aFile = new File(fileName);
		if (aFile.exists()) return aFile;
		else throw new FileNotFoundException("File  " + fileName + " does not exist.");
	}

	public static FileReader getFileReader(String fName) {
		final String OUT_FOLDER 	= "." ;

		FileReader fReader = null ;
		try {
			String inFilename = fName;
			String dirToUse = OUT_FOLDER ;
            File f = new File(dirToUse, inFilename);
            FileInputStream fiS = FileUtils.openInputStream(f) ;
            FileReader fr = new FileReader(fiS.getFD()) ;
            fReader = fr ;
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
		}
		return fReader ;
	}
}