# xls2csv

Java app to convert XLS / XLSX to CSV.

The app is simply a merge of XLS2CSV and XLSX2CSV, which is examples on Apache POI.

Really useful app to batch export an Excel files to separate CSV.

## Usage

```
java -jar xls2csv [.xls|.xlsx file] [min-column]
```

* _string_ `.xls|.xlsx file`: The input file need to convert.
* _integer_ `min-column`: Generate at least this column.
