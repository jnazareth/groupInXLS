package gpGroupXLS.format;

import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Collections;

public class Export {
    public interface XLSHeaders {
        final String H_TRANSACTION_AMOUNTS	        = "transaction amounts" ;
        final String H_OWE							= "(you owe) / owed to you" ;
        final String H_INDIVIDUAL_TOTALS	        = "individual \"spent\"" ;
        final String H_ITEM							= "Item" ;
        final String H_CATEGORY						= "Category" ;
        final String H_VENDOR						= "Vendor" ;
        final String H_DESCRIPTION					= "Description" ;
        final String H_AMOUNT						= "Amount" ;
        final String H_FROM							= "From" ;
        final String H_TO							= "To" ;
        final String H_ACTION						= "Action" ;
        final String H_CHECKSUM_TRANSACTION			= "cs(Transaction)" ;
        final String H_CHECKSUM_GROUPTOTALS        	= "cs(GroupTotals)" ;
        final String H_INDIVIDUAL_PAID		        = "individual \"paid\"" ;
        final String H_CHECKSUM_INDIVIDUALTOTALS	= "cs(IndividualTotals)" ;
    }

    public interface ExportKeys {
        final String keyItem                        = "item" ;
        final String keyCategory                    = "category" ;
        final String keyVendor                      = "vendor" ;
        final String keyDescription                 = "description" ;
        final String keyAmount                      = "amount" ;
        final String keyFrom                        = "from" ;
        final String keyTo                          = "to" ;
        final String keyAction                      = "action" ;
        final String keyTransactions                = "transactions" ;
        final String keyOwe                         = "owe" ;
        final String keyCheckSumTransaction         = "checksumTransaction" ;
        final String keySpent                       = "spent" ;
        final String keyCheckSumGroupTotals         = "checksumGroupTotals" ;
        final String keyPaid                        = "paid" ;
        final String keyCheckSumIndividualTotals    = "checksumIndividualTotals" ;
    }

    // members
    public RowLayout header0   = new RowLayout();
    public RowLayout header1   = new RowLayout();

	public void buildHeaders(String group) {
		// temp, mock data
		final int numPersons = 4 ;

        header0.empty();
        header1.empty();

        int pos = 1 ;
		header1.addCell(pos++, ExportKeys.keyItem,      XLSHeaders.H_ITEM) ;
		header1.addCell(pos++, ExportKeys.keyCategory,  XLSHeaders.H_CATEGORY) ;
		header1.addCell(pos++, ExportKeys.keyVendor,    XLSHeaders.H_VENDOR) ;
		header1.addCell(pos++, ExportKeys.keyDescription, XLSHeaders.H_DESCRIPTION) ;
		header1.addCell(pos++, ExportKeys.keyAmount,    XLSHeaders.H_AMOUNT) ;
		header1.addCell(pos++, ExportKeys.keyFrom,      XLSHeaders.H_FROM) ;
		header1.addCell(pos++, ExportKeys.keyTo,        XLSHeaders.H_TO) ;
		header1.addCell(pos++, ExportKeys.keyAction,    XLSHeaders.H_ACTION) ;

        header0.addCell(pos, ExportKeys.keyTransactions, XLSHeaders.H_TRANSACTION_AMOUNTS) ;
	    pos += numPersons ;
        header0.addCell(pos, ExportKeys.keyOwe, XLSHeaders.H_OWE) ;
	    pos += numPersons ;
        header0.addCell(pos, ExportKeys.keySpent, XLSHeaders.H_INDIVIDUAL_TOTALS) ;
	    pos += numPersons ;
        header0.addCell(pos, ExportKeys.keyPaid, XLSHeaders.H_INDIVIDUAL_PAID) ;

        //header0.dumpCollection();
    }

    /*
     * class RowLayout
     */

    public class RowLayout {
		ArrayList<CellLayout> m_Cells  = new ArrayList<CellLayout>();

        public RowLayout(RowLayout rl) {
            for (CellLayout c : rl.m_Cells) {
                CellLayout cl = rl.getCell(c.xlsPositionName) ;
                this.addCell(cl.xlsPosition, cl.xlsPositionName, cl.xlsPositionValue) ;
            }
        }

        public RowLayout() {
        }

        void addCell(int xlsPos, String xlsKey, String sValue) {
            CellLayout cl = new CellLayout(xlsPos, xlsKey, sValue) ;
            m_Cells.add(cl) ;
        }

        public CellLayout getCell (String posName) {
            for (CellLayout c : m_Cells) {
                if (c.xlsPositionName.equalsIgnoreCase(posName)) return c ;
            }
            return null ;
        }

        public CellLayout getCell (int pos) {
            for (CellLayout c : m_Cells) {
                if (c.xlsPosition == pos) return c ;
            }
            return null ;
        }

        public CellLayout setValue (int i, String v) {
            CellLayout cl = getCell(i) ;
            if (cl != null) {
                cl.xlsPositionValue = v;
                return cl ;
            }
            return null ;
        }
        public CellLayout setValue (String posName, String v) {
            CellLayout cl = getCell(posName) ;
            if (cl != null) {
                cl.xlsPositionValue = v;
                return cl ;
            }
            return null ;
        }

        public int length() {
            return m_Cells.size();
        }

        public void empty(){
            m_Cells.clear();
        }

        public void dumpCollection() {
            for (CellLayout c : m_Cells) {
                System.out.println(c.xlsPosition + "|" + c.xlsPositionName + "|" + c.xlsPositionValue);
            }
        }

        public class CellLayout {
            public int		xlsPosition;
            public String   xlsPositionName;
            public String	xlsPositionValue;

            public CellLayout(int x, String n, String v) {
                xlsPosition = x;
                xlsPositionName = n;
                xlsPositionValue = v;
            }
        }
    }
}
