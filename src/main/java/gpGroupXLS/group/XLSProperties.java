package gpGroupXLS.group;

import gpGroupXLS.format.Export;

public class XLSProperties {
    public static final String SUMMARY_SHEET_NAME = "Summary";
    public static final String XRATES_SHEET_NAME = "XRates";
    public static final String TOTALS_LABEL = "Totals:";

    public static final Export HEADER_EXPORT = new Export();

    // Target format
    public static final int TOTAL_COLUMN_POSITION = 0; // Position of "total" column
    public static final int GROUP_NAME_COLUMN_OFFSET = 1; // Offset of "group name" column

    public static final int ZERO_COLUMN_OFFSET = 1; // POI col index vs. JSON spec.
    public static int numberToSkip = 0;

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("XLSProperties{");
        sb.append("SUMMARY_SHEET_NAME='").append(SUMMARY_SHEET_NAME).append('\'');
        sb.append(", XRATES_SHEET_NAME='").append(XRATES_SHEET_NAME).append('\'');
        sb.append(", TOTALS_LABEL='").append(TOTALS_LABEL).append('\'');
        sb.append(", TOTAL_COLUMN_POSITION=").append(TOTAL_COLUMN_POSITION);
        sb.append(", GROUP_NAME_COLUMN_OFFSET=").append(GROUP_NAME_COLUMN_OFFSET);
        sb.append(", ZERO_COLUMN_OFFSET=").append(ZERO_COLUMN_OFFSET);
        sb.append(", numberToSkip=").append(numberToSkip);
        sb.append('}');
        return sb.toString();
    }
}