package com.xlscsv.converter;

import org.apache.poi.openxml4j.opc.*;

import java.io.File;

public class Converter {
    public static void convert(String fileName, String outputPath, int minColumns, String... sheetName) throws Exception {
        File inputFile = new File(fileName);
        String fileNameNoEx = getFileNameNoEx(inputFile.getName());
        if (!inputFile.exists()) {
            System.err.println("Not found or not a file: " + inputFile.getPath());
            return;
        }

        OutputDispatcher dispatcher;

        if (fileName.endsWith("xls") || fileName.endsWith("xlsx")) {
            dispatcher = new PerSheetOutputDispatcher(outputPath + fileNameNoEx);
        } else {
            dispatcher = new StandardOutputDispatcher();
        }

//        int minColumns = -1;

        if (fileName.endsWith("xlsx")) {
            // The package open is instantaneous, as it should be.
            OPCPackage p = OPCPackage.open(inputFile.getPath(), PackageAccess.READ);
            XlsxToCsv xlsx2csv = new XlsxToCsv(p, dispatcher, minColumns);
            xlsx2csv.process();

        } else if (fileName.endsWith("xls")) {
            XlsToCsv xls2csv = new XlsToCsv(fileName, dispatcher, minColumns);
            xls2csv.process();
        } else if (fileName.endsWith("csv")){
            String sn;
            if(sheetName != null && sheetName.length == 1){
                sn = sheetName[0];
            }else{
                sn = "Sheet1";
            }
//            new CsvToXls().process(fileName, outputPath + fileNameNoEx, sn);
            new CsvToXls().process(fileName, outputPath, sn);
        }
        else {
            throw new Exception("Unrecognized file type");
        }
    }

    public static String getFileNameNoEx(String filename) {
        if ((filename != null) && (filename.length() > 0)) {
            int dot = filename.lastIndexOf('.');
            if ((dot >-1) && (dot < (filename.length()))) {
                return filename.substring(0, dot);
            }
        }
        return filename;
    }
}