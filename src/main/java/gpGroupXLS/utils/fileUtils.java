package gpGroupXLS.utils;

import java.io.File;
import java.io.FileNotFoundException;

public class fileUtils {
	public static File getFile(String fileName)
	throws FileNotFoundException
	{
		File aFile = new File(fileName);
		if (aFile.exists()) return aFile;
		else throw new FileNotFoundException("File  " + fileName + " does not exist.");
	}
}