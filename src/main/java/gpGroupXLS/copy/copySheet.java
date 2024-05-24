package gpGroupXLS.copy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gpGroupXLS.group.XLSProperties;
import gpGroupXLS.group.tabGroup;
import gpGroupXLS.tabs.TabSummary2;
import gpGroupXLS.tabs.TabSummary2.tabEntry2;
import gpGroupXLS.tabs.TabSummary2.tabGroupBase;

public class copySheet {
    public copySheet() {
        // constructor
    }

    private void copyCell(XSSFWorkbook wbSummary, XSSFWorkbook wbToBeCopied, Cell currentCell, Cell newCell) {
        try {
            XSSFCellStyle cellStyle = wbSummary.createCellStyle();
            DataFormat dFormat = wbSummary.createDataFormat();
            String formatStr = currentCell.getCellStyle().getDataFormatString();
            cellStyle.setDataFormat(dFormat.getFormat(formatStr));
    
            XSSFFont oldFont = wbToBeCopied.getFontAt(currentCell.getCellStyle().getFontIndex());
            XSSFFont newFont = wbSummary.createFont();
            newFont.setBold(oldFont.getBold());
            newFont.setColor(oldFont.getColor());
            newFont.setFontHeight(oldFont.getFontHeight());
            newFont.setFontName(oldFont.getFontName());
            newFont.setItalic(oldFont.getItalic());
            newFont.setStrikeout(oldFont.getStrikeout());
            newFont.setTypeOffset(oldFont.getTypeOffset());
            newFont.setUnderline(oldFont.getUnderline());
            newFont.setCharSet(oldFont.getCharSet());
            cellStyle.setFont(newFont);

            cellStyle.setAlignment(currentCell.getCellStyle().getAlignment());
            cellStyle.setHidden(currentCell.getCellStyle().getHidden());
            cellStyle.setLocked(currentCell.getCellStyle().getLocked());
            cellStyle.setWrapText(currentCell.getCellStyle().getWrapText());
            cellStyle.setBorderBottom(currentCell.getCellStyle().getBorderBottom());
            cellStyle.setBorderLeft(currentCell.getCellStyle().getBorderLeft());
            cellStyle.setBorderRight(currentCell.getCellStyle().getBorderRight());
            cellStyle.setBorderTop(currentCell.getCellStyle().getBorderTop());
            cellStyle.setFillBackgroundColor(currentCell.getCellStyle().getFillBackgroundColor());
            cellStyle.setFillForegroundColor(currentCell.getCellStyle().getFillForegroundColor());
            cellStyle.setFillPattern(currentCell.getCellStyle().getFillPattern());
            cellStyle.setIndention(currentCell.getCellStyle().getIndention());
            cellStyle.setBottomBorderColor(currentCell.getCellStyle().getBottomBorderColor());
            cellStyle.setLeftBorderColor(currentCell.getCellStyle().getLeftBorderColor());
            cellStyle.setRightBorderColor(currentCell.getCellStyle().getRightBorderColor());
            cellStyle.setTopBorderColor(currentCell.getCellStyle().getTopBorderColor());
            cellStyle.setRotation(currentCell.getCellStyle().getRotation());
            cellStyle.setVerticalAlignment(currentCell.getCellStyle().getVerticalAlignment());
            newCell.setCellStyle(cellStyle);

            switch(currentCell.getCellType()) {
                case STRING:
                    String cV = currentCell.getStringCellValue() ;
                    newCell.setCellValue(cV);
                    break;
                case NUMERIC:
                    Double d = currentCell.getNumericCellValue() ;
                    newCell.setCellValue(d);
                    break;
                default:
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

	private int copyContents2(XSSFWorkbook wbSummary, XSSFSheet sheetTarget, String groupName, XSSFSheet sheetToBeCopied) {
        try {
            int rowAdded = -1 ;
            int row = sheetToBeCopied.getLastRowNum() ;

            XSSFWorkbook wbToBeCopied = sheetToBeCopied.getWorkbook();
            for (int r = 0; r < row; r++) {
                Row rowitr = sheetToBeCopied.getRow(r) ;
                if (rowitr != null) {
                    Row currentRow = sheetTarget.createRow(r);
                    rowAdded = currentRow.getRowNum();

                    int lastColumn = rowitr.getLastCellNum() ;
                    for (int col = 0; col < lastColumn; col++) {
                        Cell celldata = rowitr.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (celldata == null) {
                            // The spreadsheet is empty in this cell
                        } else {
                            Cell cellTarget = currentRow.createCell(col);
                            copyCell(wbSummary, wbToBeCopied, celldata, cellTarget) ;
                        }
                    }
                }
            }
            return (rowAdded) ;
        } catch (Exception e) {
            e.printStackTrace();
            return -1;
        }
    }   
    
    private String makeSheetName (String fileName, String groupName) {
        final String _sep = "." ;
        String p = "";
        int nPrefix = fileName.indexOf(_sep) ;
        if (nPrefix != -1) p = fileName.substring(0, nPrefix + 1);

        String s = ""; 
        int nSuffix = groupName.lastIndexOf(_sep) ;
        int nSuffix2 = 0;
        if (nSuffix != -1) nSuffix2 = groupName.lastIndexOf(_sep, nSuffix-1);
        if (nSuffix2 != -1) s = groupName.substring(nSuffix2 + 1);

        String n = fileName + _sep + groupName;
        int h = Math.abs(n.hashCode() % 10000) ;    // last 4 digits
        String hashName = p + "xlsx." + s + _sep + h ;
        //System.out.println(groupName + "\t" + hashName + "\t" + hashName.length()) ;

        return hashName;
    }

	public void buildCopySheets(XSSFWorkbook workBookGroup, tabGroup tg) {
		try {
			XSSFWorkbook workBookIn = null;
			TabSummary2 ts2 = tg.m_tabSummary ;
			for (tabGroupBase tgb : ts2.m_groupTabs) {
				tabEntry2 gItem = tgb.te;
				File fIn = new File(gItem.fileName);
				FileInputStream fileIn = new FileInputStream(fIn);
				workBookIn = new XSSFWorkbook(fileIn);

				XSSFCreationHelper createHelper = workBookGroup.getCreationHelper();
				XSSFHyperlink sheetLink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);

				copySheet cSheet = new copySheet() ;
				String sName = cSheet.copyASheet(workBookIn, gItem.fileName, gItem.groupName, workBookGroup) ;
				if (sName != null) {
					int p = tgb.rowNumber;
					if (p != -1) {
						XSSFSheet sheetSummary = workBookGroup.getSheet(XLSProperties._SummarySheetName) ;
						if (sheetSummary == null) break ;

						Row r = sheetSummary.getRow(p) ;
						if (r == null) break ;

						Cell cellSheetReference = r.createCell(0);
						if (cellSheetReference == null) break ;

						String sNameR = sName + "!$A$1" ;
						CellReference s2a1 = new CellReference(sNameR);
						if (s2a1 != null) {
							String sLink = s2a1.formatAsString();
							cellSheetReference.setCellValue(gItem.groupName) ;//sLink
							sheetLink.setAddress(sLink);
							cellSheetReference.setHyperlink(sheetLink);
							//cellSheetReference.setCellStyle(hlink_style);
						}
					}
				}
				fileIn.close() ;
			}
			try {
				workBookIn.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

    public String copyASheet(XSSFWorkbook workBookIn, String fileName, String groupName, XSSFWorkbook workBookGroup) {
        try {
            XSSFSheet sheetToBeCopied = workBookIn.getSheet(groupName) ;
            if (sheetToBeCopied == null) return null ;

            String tName = makeSheetName(fileName, groupName);
            XSSFSheet sheetTarget = workBookGroup.createSheet(tName);
            if (sheetTarget == null) return null ;

            int rowsCopied = copyContents2(workBookGroup, sheetTarget, groupName, sheetToBeCopied) ;
            return tName;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    @Override
    public String toString() {
        // TODO Auto-generated method stub
        return super.toString();
    }    
}
