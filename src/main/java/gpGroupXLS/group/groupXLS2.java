package gpGroupXLS.group;

import gpGroupXLS.copy.copySheet;
import gpGroupXLS.format.Export;
import gpGroupXLS.json.ReadJson;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class groupXLS2 {
	XRateGrid xg = new XRateGrid() ;
	BaseGrid bg = new BaseGrid() ;
	TargetGrid tg2 = new TargetGrid() ;
	copySheet cs2 = new copySheet() ;

	public File InitializeXLS(String fName) {
        File f = null ;
		try {
            String xlsFile = fName;

			f = new File(xlsFile);
            String filetoRecreate = f.getName();
            boolean dFile = false;
			if (f.exists()) {
                dFile = f.delete();
				if (!dFile) System.out.println("failed to delete:" + filetoRecreate);
            }
			File f2 = new File(filetoRecreate);
			f = f2;
			XSSFWorkbook workBook = new XSSFWorkbook();
			try (FileOutputStream fileOut = new FileOutputStream(f)) {
				workBook.write(fileOut);
				workBook.close();
				fileOut.close() ;
			}

			// _newFormatting 4/22
			XLSProperties.eHeader.buildHeaders("");
			XLSProperties._numberToSkip = (XLSProperties.eHeader.header0.getCell(Export.ExportKeys.keyOwe).xlsPosition) - (XLSProperties.zeroColOffset);
		
			return f;
        } catch (FileNotFoundException fnf) {
        } catch (IOException ioe) {
        }
        return f;
	}

	private boolean buildGroupXLS2(File hXLS, tabGroup tg) {
		try {
			File fGroup = new File(hXLS.getName());
			FileInputStream fileGroup = new FileInputStream(fGroup);
			XSSFWorkbook workBookGroup = new XSSFWorkbook(fileGroup);

			xg.buildXRateGrid(workBookGroup, tg) ;
			bg.buildBaseGrid(workBookGroup, tg) ;
			tg2.buildTargetGrid(workBookGroup, tg) ;
			cs2.buildCopySheets(workBookGroup, tg) ;

			try {
				FileOutputStream outputStream = new FileOutputStream(hXLS.getName());
				workBookGroup.write(outputStream);
				workBookGroup.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			fileGroup.close() ;

			return true ;
		} catch (FileNotFoundException nfe) {
			nfe.printStackTrace();
			return false ;
		} catch (Exception e) {
			e.printStackTrace();
			return false ;
		}
	}

	private tabGroup readFromJSON(String configJSON, String xlsGroupFile) {
		try {
			ReadJson rj = new ReadJson() ;
			tabGroup tg = rj.readJSONConfigFile2(configJSON, xlsGroupFile);
			return tg ;
		} catch (Exception e) {
			System.err.println("readFromJSON::Exception::" + e.getMessage()) ;
			return null;
		}
	}

	public void ReadXLSBuildGroup2(String configJSON, String xlsGroupFile) {
		try {
			tabGroup tg = readFromJSON(configJSON, xlsGroupFile);
			if (tg == null) return ;

			File hFile = InitializeXLS(xlsGroupFile);
			if (hFile == null) return ;

			boolean b = buildGroupXLS2(hFile, tg);
		} catch (Exception e) {
			System.err.println("ReadXLSBuildGroup::Exception: " + e.getMessage());
		}
	}
}