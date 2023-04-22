package gpGroupXLS.tabs;

import gpGroupXLS.xls._Coordinates;

import java.util.ArrayList;

public class TabSummary2 {
	public ArrayList<tabGroupBase>	m_groupTabs = null;
		String xlsFileName;
		public _Coordinates coords;

	public void setXLSFileName(String xFN) {
		xlsFileName = xFN;
	}
	public String getXLSFileName() {
		return xlsFileName ;
	}

	public _Coordinates getCoords() {
		return coords;
	}

	public void setCoords(String cd) {
		coords = new _Coordinates(cd);
	}

	public ArrayList<tabGroupBase> addItem(String fN, String gN, String c, String f) {
		tabEntry2 tE = new tabEntry2(fN, gN, c, f) ;
		tabGroupBase tgb = new tabGroupBase(tE);
		if (m_groupTabs == null) m_groupTabs = new ArrayList<tabGroupBase>() ;
		boolean b = m_groupTabs.add(tgb) ;

		return m_groupTabs ;
	}

	public void dump() {
		for (tabGroupBase tgb : m_groupTabs) {
			tgb.te.dump() ;
		}
	}

	public class tabGroupBase {
		public tabEntry2	te;	// read from input
		public int	rowNumber;	// populate once added to XLS

		public tabGroupBase(tabEntry2 t){
			te = t ;
			rowNumber = -1 ;
		}
	}

	public class tabEntry2 {
		public String fileName;
		public String groupName;
		public String currency;
		public String format;

		public tabEntry2() {
		}

		public tabEntry2(String fN, String gN, String c, String f) {
			fileName = fN;
			groupName = gN;
			currency = c;
			format = f;
		}

		public void dump() {
			final String _SEP = "|" ;
			System.out.println(fileName + _SEP + groupName + _SEP + currency + _SEP + format + _SEP + coords.toCoordsString()) ;
		}
	}
}
