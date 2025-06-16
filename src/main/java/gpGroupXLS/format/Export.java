package gpGroupXLS.format;

import java.util.ArrayList;
import java.util.List;

public class Export {

    public interface XLSHeaders {
        String TRANSACTION_AMOUNTS = "transaction amounts";
        String OWE = "(you owe) / owed to you";
        String INDIVIDUAL_TOTALS = "individual \"spent\"";
        String ITEM = "Item";
        String CATEGORY = "Category";
        String VENDOR = "Vendor";
        String DESCRIPTION = "Description";
        String AMOUNT = "Amount";
        String FROM = "From";
        String TO = "To";
        String ACTION = "Action";
        String CHECKSUM_TRANSACTION = "cs(Transaction)";
        String CHECKSUM_GROUPTOTALS = "cs(GroupTotals)";
        String INDIVIDUAL_PAID = "individual \"paid\"";
        String CHECKSUM_INDIVIDUALTOTALS = "cs(IndividualTotals)";
    }

    public interface ExportKeys {
        String ITEM = "item";
        String CATEGORY = "category";
        String VENDOR = "vendor";
        String DESCRIPTION = "description";
        String AMOUNT = "amount";
        String FROM = "from";
        String TO = "to";
        String ACTION = "action";
        String TRANSACTIONS = "transactions";
        String OWE = "owe";
        String CHECKSUM_TRANSACTION = "checksumTransaction";
        String SPENT = "spent";
        String CHECKSUM_GROUPTOTALS = "checksumGroupTotals";
        String PAID = "paid";
        String CHECKSUM_INDIVIDUALTOTALS = "checksumIndividualTotals";
    }

    public final RowLayout header0 = new RowLayout();
    private final RowLayout header1 = new RowLayout();

    public void buildHeaders(int numPersons) {
        header0.clear();
        header1.clear();

        int position = 1;
        header1.addCell(position++, ExportKeys.ITEM, XLSHeaders.ITEM);
        header1.addCell(position++, ExportKeys.CATEGORY, XLSHeaders.CATEGORY);
        header1.addCell(position++, ExportKeys.VENDOR, XLSHeaders.VENDOR);
        header1.addCell(position++, ExportKeys.DESCRIPTION, XLSHeaders.DESCRIPTION);
        header1.addCell(position++, ExportKeys.AMOUNT, XLSHeaders.AMOUNT);
        header1.addCell(position++, ExportKeys.FROM, XLSHeaders.FROM);
        header1.addCell(position++, ExportKeys.TO, XLSHeaders.TO);
        header1.addCell(position++, ExportKeys.ACTION, XLSHeaders.ACTION);

        header0.addCell(position, ExportKeys.TRANSACTIONS, XLSHeaders.TRANSACTION_AMOUNTS);
        position += numPersons;
        header0.addCell(position, ExportKeys.OWE, XLSHeaders.OWE);
        position += numPersons;
        header0.addCell(position, ExportKeys.SPENT, XLSHeaders.INDIVIDUAL_TOTALS);
        position += numPersons;
        header0.addCell(position, ExportKeys.PAID, XLSHeaders.INDIVIDUAL_PAID);
    }

    public class RowLayout {
        private final List<CellLayout> cells = new ArrayList<>();

        public RowLayout() {
        }

        public RowLayout(RowLayout other) {
            for (CellLayout cell : other.cells) {
                addCell(cell.position, cell.key, cell.value);
            }
        }

        public void addCell(int position, String key, String value) {
            cells.add(new CellLayout(position, key, value));
        }

        public CellLayout getCell(String key) {
            return cells.stream()
                        .filter(cell -> cell.key.equalsIgnoreCase(key))
                        .findFirst()
                        .orElse(null);
        }

        public CellLayout getCell(int position) {
            return cells.stream()
                        .filter(cell -> cell.position == position)
                        .findFirst()
                        .orElse(null);
        }

        public CellLayout setValue(int position, String value) {
            CellLayout cell = getCell(position);
            if (cell != null) {
                cell.value = value;
            }
            return cell;
        }

        public CellLayout setValue(String key, String value) {
            CellLayout cell = getCell(key);
            if (cell != null) {
                cell.value = value;
            }
            return cell;
        }

        public int size() {
            return cells.size();
        }

        public void clear() {
            cells.clear();
        }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder();
            for (CellLayout cell : cells) {
                sb.append(cell).append("\n");
            }
            return sb.toString();
        }

        public class CellLayout {
            public final int position;
            private final String key;
            private String value;

            public CellLayout(int position, String key, String value) {
                this.position = position;
                this.key = key;
                this.value = value;
            }

            @Override
            public String toString() {
                StringBuilder sb = new StringBuilder();
                sb.append(position).append("|").append(key).append("|").append(value);
                return sb.toString();
            }
        }
    }
}