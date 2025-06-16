package gpGroupXLS.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.commons.io.FileUtils;

public class fileUtils {

    private static final String DEFAULT_DIRECTORY = ".";

    public static File getFile(String fileName) throws FileNotFoundException {
        File file = new File(fileName);
        if (file.exists()) {
            return file;
        } else {
            throw new FileNotFoundException("File " + fileName + " does not exist.");
        }
    }

    public static FileReader getFileReader(String fileName) {
        try {
            File file = new File(DEFAULT_DIRECTORY, fileName);
            FileInputStream fileInputStream = FileUtils.openInputStream(file);
            return new FileReader(fileInputStream.getFD());
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    @Override
    public String toString() {
        return "FileUtils{}";
    }
}