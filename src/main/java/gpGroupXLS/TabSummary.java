package gpGroupXLS;

import java.util.LinkedHashSet;

public class TabSummary {
	public LinkedHashSet<tabEntry> hsTabSummary = new LinkedHashSet<tabEntry>();
	String xlsFileName;

	public void setXLSFileName(String xFN) {
		xlsFileName = xFN;
	}
	public String getXLSFileName() {
		return xlsFileName ;
	}

	LinkedHashSet<tabEntry> addItem(String fN, String gN, Double r, String c, String f, String xc) {
		tabEntry tE = new tabEntry(fN, gN, r, c, f, xc) ;
		hsTabSummary.add(tE) ;
		return hsTabSummary ;
	}

	public int size() {
		return hsTabSummary.size();
	}

	public void dump() {
		for (tabEntry g : this.hsTabSummary) {
			g.dump();
		}
	}

	public class tabEntry {
		String fileName;
		String groupName;
		public Double rate;
		public String currency;
		String format;
		public String xcurrency;

		public tabEntry() {
		}

		public tabEntry(String fN, String gN, Double r, String c, String f, String xc) {
			fileName = fN;
			groupName = gN;
			rate = r;
			currency = c;
			format = f;
			xcurrency = xc;
		}

		public void dump() {
			final String _SEP = "|" ;
			System.out.println(fileName + _SEP + groupName + _SEP + rate + _SEP + currency + _SEP + format + _SEP + xcurrency) ;
		}
	}
}
