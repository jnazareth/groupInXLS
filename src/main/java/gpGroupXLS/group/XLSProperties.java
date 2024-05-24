package gpGroupXLS.group;

import gpGroupXLS.format.Export;
import gpGroupXLS.format.Export.RowLayout;

public class XLSProperties {
	public static final String _SummarySheetName 	= "Summary" ;
	public static final String _XRatesSheetName 	= "XRates" ;
	public static final String _Totals 				= "Totals:" ;

	public static Export eHeader = new Export() ;

	// target format
	public static final int _totalColumnPosition 	= 0 ;	// position of "total" column
	public static final int	_groupNameColOffset 	= 1 ;	// offset of "group name" column

	public final static int zeroColOffset 			= 1;	// POI col index vs. JSON spec.
	//private final static RowLayout.CellLayout rcl = eHeader.header0.getCell(Export.ExportKeys.keyOwe) ;
	public static int _numberToSkip = 0;//rcl.xlsPosition - zeroColOffset;
}