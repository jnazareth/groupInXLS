package gpGroupXLS.xchg;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;

import gpGroupXLS.currency.CurrencyInfo;

public class ExchangeRateTable {
    private LinkedHashMap<String, TargetCurrencies> targetGrid = new LinkedHashMap<>();
    private String sheetName = "";
    private List<Integer> rowRefIndex = new ArrayList<>();
    private List<Integer> colRefIndex = new ArrayList<>();

    public void addRates(String[] fromCurrencies, String toCurrency, String format, RateDate[] rateDates) {
        TargetCurrencies targetCurrencies = new TargetCurrencies(fromCurrencies, toCurrency, format, rateDates);
        targetGrid.put(toCurrency, targetCurrencies);
    }

    public void buildRateReferenceGrid(CellReference cellReference) {
        sheetName = cellReference.getCellRefParts()[0];
        int startRowReference = Integer.parseInt(cellReference.getCellRefParts()[1]);
        char startColReference = cellReference.getCellRefParts()[2].charAt(0);

        int row = startRowReference;
        char col = startColReference;

        boolean rowsAdded = false;
        for (Map.Entry<String, TargetCurrencies> entry : targetGrid.entrySet()) {
            TargetCurrencies currencies = entry.getValue();
            row = startRowReference;
            col = (char) (col + 1);
            int colIndex = col;

            colRefIndex.add(colIndex);
            if (!rowsAdded) {
                for (@SuppressWarnings("unused") ExchangePair pair : currencies.getTargetRates()) {
                    rowRefIndex.add(++row);
                }
                rowsAdded = true;
            }
            col = (char) (col + 1);
        }
    }

    public String getRateReference(int fromCurrencyIndex, int toCurrencyIndex) {
        int rowIndex = rowRefIndex.get(fromCurrencyIndex);
        int colIndex = colRefIndex.get(toCurrencyIndex);
        char colChar = (char) colIndex;
        return sheetName + "!" + colChar + rowIndex;
    }

	public LinkedHashMap<String, TargetCurrencies> getTargetGrid() {
        return targetGrid;
    }

    @SuppressWarnings("unused")
    public void printExchangeRates() {
        int rowIndex = 0;
        for (Map.Entry<String, TargetCurrencies> entry : targetGrid.entrySet()) {
            TargetCurrencies currencies = entry.getValue();
            int colIndex = 0;
            for (ExchangePair pair : currencies.getTargetRates()) {
                String rateReference = getRateReference(colIndex, rowIndex);
                System.out.println("Rate Reference: " + rateReference);
                colIndex++;
            }
            rowIndex++;
        }
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder("ExchangeRateTable{");
        sb.append("sheetName='").append(sheetName).append('\'');
        sb.append(", targetGrid=").append(targetGrid);
        sb.append('}');
        return sb.toString();
    }

    public class TargetCurrencies {
        private List<ExchangePair> targetRates = new ArrayList<>();
        private CurrencyInfo currencyInfo;

        public TargetCurrencies(String[] fromCurrencies, String toCurrency, String format, RateDate[] rateDates) {
            if (fromCurrencies.length != rateDates.length) return;
            for (int i = 0; i < rateDates.length; i++) {
                ExchangePair pair = new ExchangePair(fromCurrencies[i], toCurrency, rateDates[i].rate, rateDates[i].date);
                targetRates.add(pair);
            }
            currencyInfo = new CurrencyInfo(toCurrency, format);
        }

        public List<ExchangePair> getTargetRates() {
            return targetRates;
        }

		public CurrencyInfo getCurrencyInfo() {
			return currencyInfo;
		}

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder("TargetCurrencies{");
            sb.append("targetRates=").append(targetRates);
            sb.append(", currencyInfo=").append(currencyInfo);
            sb.append('}');
            return sb.toString();
        }
    }

    public class ExchangePair {
        private String fromCurrency;
        private String toCurrency;
        private RateDate rateDate;

        public ExchangePair(String fromCurrency, String toCurrency, Double rate, String date) {
            this.fromCurrency = fromCurrency;
            this.toCurrency = toCurrency;
            this.rateDate = new RateDate(rate, date);
        }

		public String getFromCurrency() {
			return fromCurrency;
		}

		public String getToCurrency() {
			return toCurrency;
		}

		public RateDate getRateDate() {
			return rateDate;
		}

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder("ExchangePair{");
            sb.append("fromCurrency='").append(fromCurrency).append('\'');
            sb.append(", toCurrency='").append(toCurrency).append('\'');
            sb.append(", rateDate=").append(rateDate);
            sb.append('}');
            return sb.toString();
        }
    }

    public class RateDate {
        private Double rate;
        private String date;

        public RateDate(Double rate, String date) {
            this.rate = rate;
            this.date = date;
        }

		public Double getRate() {
 		  return rate;
		}

		public String getDate() {
			return date;
		}

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder("RateDate{");
            sb.append("rate=").append(rate);
            sb.append(", date='").append(date).append('\'');
            sb.append('}');
            return sb.toString();
        }
    }
}