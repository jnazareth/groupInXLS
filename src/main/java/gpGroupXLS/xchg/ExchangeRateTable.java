package gpGroupXLS.xchg;

import gpGroupXLS.TabSummary;
import gpGroupXLS.TabSummary.tabEntry;

import java.util.ArrayList;

public class ExchangeRateTable {
    ArrayList<String> FromToString = null ;
    Double[][] exchangeRateTable = null ; 

    public void ExchangeRateTable() {
        // constructor
    }

    public void buildExchangeRateTable(TabSummary ts) {
        Double[][] exchangeRateTable = new Double[ts.size()][ts.size()];
        ArrayList<String> FromString = new ArrayList<String>(ts.size()) ;
        ArrayList<String> ToString = new ArrayList<String>(ts.size()) ;

        System.out.println("ts.size():" + ts.size());
        
        for (tabEntry gItem : ts.hsTabSummary) {
            System.out.println("gItem.currency:" + gItem.currency);
            FromString.add(gItem.currency) ;
            ToString.add(gItem.xcurrency) ;
        }
        for (int k = 0; k < FromString.size(); k++) {
            System.out.print(FromString.get(k) + "\t");
        }
        System.out.println("");
        for (int k = 0; k < ToString.size(); k++) {
            System.out.print(ToString.get(k) + "\t");
        }
        System.out.println("");

        for (tabEntry gItem : ts.hsTabSummary) {
            System.out.println("gItem.currency:" + gItem.currency + ", gItem.xcurrency:" + gItem.xcurrency);
            int f = FromString.indexOf(gItem.currency);
            int t = ToString.indexOf(gItem.xcurrency);

            System.out.println("f:" + f + ", t:" + t);
            exchangeRateTable[f][t] = gItem.rate ;   
        }
        System.out.println("exchangeRateTable.length:" + exchangeRateTable.length);

        for (int i = 0; i < exchangeRateTable.length; i++) {
            for (int j = 0; j < exchangeRateTable.length; j++) {
                System.out.print("[" + i + "][" + j + "]:" + exchangeRateTable[i][j] + "\t");
            }
            System.out.println("");
        }
    }
}
