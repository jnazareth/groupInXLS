package gpGroupXLS.tabs;

import gpGroupXLS.xls._Coordinates;

import java.util.ArrayList;
import java.util.List;

public class TabSummary2 {
    private List<TabGroupBase> groupTabs = new ArrayList<>();
    private String xlsFileName;
    private _Coordinates coords;
    private int numPersons;

    public void setXLSFileName(String xlsFileName) {
        this.xlsFileName = xlsFileName;
    }

    public String getXLSFileName() {
        return xlsFileName;
    }

    public _Coordinates getCoords() {
        return coords;
    }

    public void setCoords(String coordsString) {
        this.coords = new _Coordinates(coordsString);
    }

    public void setNumPersons(int numPersons) {
        this.numPersons = numPersons;
    }

    public int getNumPersons() {
        return numPersons;
    }

    public List<TabGroupBase> addItem(String fileName, String groupName, String currency, String format) {
        TabEntry2 tabEntry = new TabEntry2(fileName, groupName, currency, format);
        TabGroupBase tabGroupBase = new TabGroupBase(tabEntry);
        groupTabs.add(tabGroupBase);
        return groupTabs;
    }

    public List<TabGroupBase> getGroupTabs() {
        return groupTabs;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder("TabSummary2{");
        sb.append("xlsFileName='").append(xlsFileName).append('\'');
        sb.append(", coords=").append(coords.toCoordsString());
        sb.append(", numPersons=").append(numPersons);
        sb.append(", groupTabs=").append(groupTabs);
        sb.append('}');
        return sb.toString();
    }

    public class TabGroupBase {
        private final TabEntry2 tabEntry;
        private int rowNumber = -1;

        public TabGroupBase(TabEntry2 tabEntry) {
            this.tabEntry = tabEntry;
        }

        public TabEntry2 getTabEntry() {
            return tabEntry;
        }

        public int getRowNumber() {
            return rowNumber;
        }

        public void setRowNumber(int rowNumber) {
            this.rowNumber = rowNumber;
        }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder("TabGroupBase{");
            sb.append("tabEntry=").append(tabEntry);
            sb.append(", rowNumber=").append(rowNumber);
            sb.append('}');
            return sb.toString();
        }
    }

    public class TabEntry2 {
        private final String fileName;
        private final String groupName;
        private final String currency;
        private final String format;

        public TabEntry2(String fileName, String groupName, String currency, String format) {
            this.fileName = fileName;
            this.groupName = groupName;
            this.currency = currency;
            this.format = format;
        }

        public String getFileName() {
            return fileName;
        }

        public String getGroupName() {
            return groupName;
        }

        public String getCurrency() {
            return currency;
        }

        public String getFormat() {
            return format;
        }

        @Override
        public String toString() {
            final String separator = "|";
            StringBuilder sb = new StringBuilder("[");
            sb.append(fileName).append(separator)
              .append(groupName).append(separator)
              .append(currency).append(separator)
              .append(format).append(separator)
              .append(coords.toCoordsString())
              .append("]");
            return sb.toString();
        }
    }
}