package gpGroupXLS.group;

import gpGroupXLS.copy.CopySheet;
import gpGroupXLS.format.Export;
import gpGroupXLS.json.ReadJson;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GroupXLS2 {
    private final XRateGrid xRateGrid = new XRateGrid();
    private final BaseGrid baseGrid = new BaseGrid();
    private final TargetGrid targetGrid = new TargetGrid();
    private final CopySheet copySheet = new CopySheet();

    public File initializeXLS(String fileName, TabGroup tabGroup) {
        File file = new File(fileName);
        try {
            if (file.exists() && !file.delete()) {
                System.out.println("Failed to delete: " + file.getName());
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook();
                 FileOutputStream fileOut = new FileOutputStream(file)) {
                workbook.write(fileOut);
            }

            XLSProperties.HEADER_EXPORT.buildHeaders(tabGroup.getTabSummary().getNumPersons());
            XLSProperties.numberToSkip = XLSProperties.HEADER_EXPORT.header0.getCell(Export.ExportKeys.OWE).position - XLSProperties.ZERO_COLUMN_OFFSET;

            return file;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private boolean buildGroupXLS(File xlsFile, TabGroup tabGroup) {
        try (FileInputStream fileGroup = new FileInputStream(xlsFile);
             XSSFWorkbook workbookGroup = new XSSFWorkbook(fileGroup)) {

            xRateGrid.buildExchangeRateGrid(workbookGroup, tabGroup);
            baseGrid.buildBaseGrid(workbookGroup, tabGroup);
            targetGrid.buildTargetGrid(workbookGroup, tabGroup);
            copySheet.buildCopySheets(workbookGroup, tabGroup);

            try (FileOutputStream outputStream = new FileOutputStream(xlsFile)) {
                workbookGroup.write(outputStream);
            }

            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    private TabGroup readFromJSON(String configJSON, String xlsGroupFile) {
        try {
            ReadJson readJson = new ReadJson();
            return readJson.readJSONConfigFile(configJSON, xlsGroupFile);
        } catch (Exception e) {
            System.err.println("readFromJSON::Exception::" + e.getMessage());
            return null;
        }
    }

    public void readXLSBuildGroup(String configJSON, String xlsGroupFile) {
        try {
            TabGroup tabGroup = readFromJSON(configJSON, xlsGroupFile);
            if (tabGroup == null) return;

            File xlsFile = initializeXLS(xlsGroupFile, tabGroup);
            if (xlsFile == null) return;

            buildGroupXLS(xlsFile, tabGroup);
        } catch (Exception e) {
            System.err.println("readXLSBuildGroup::Exception: " + e.getMessage());
        }
    }
}