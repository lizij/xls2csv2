import com.xlscsv.converter.Converter;

import java.io.File;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TestMain {
    public static void main(String[] args) throws Exception{
        if (args.length < 2) {
            System.err.println("Use:");
            System.err.println("xlscsv <xls/xlsx file path> <csv directory path> [min columns]");
            System.err.println("xlscsv <csv file path> <xls file path> [min columns] [sheetName]");
            System.err.println("Note that if the file has multiple sheets each will be placed in a separate file");
            return;
        }

        String inputFileName = args[0];
        File inputFile = new File(inputFileName);
        if (!inputFile.exists()) {
            System.err.println("Not found or not a file: " + inputFile.getPath());
            return;
        }

        String outputFileName = args[1];

        int minColumns = -1;
        if (args.length >= 3)
            minColumns = Integer.parseInt(args[2]);

        String sheetName = "Sheet1";
        if(args.length >= 4)
            sheetName = args[3];

        System.out.println("Start converting");
        Converter.convert(inputFileName, outputFileName, minColumns, sheetName);
//        convertAll();
    }
    public static void convertAll() throws Exception{
        String basePath = System.getProperty("user.dir");
        String inputPath = basePath + "\\input\\";
        String outputPath = basePath + "\\output\\";

        File inputDir = new File(inputPath);
        File[] inputFiles = inputDir.listFiles();
        if(inputFiles != null){
            for(File tmpFile:inputFiles){
                if(tmpFile.getName().endsWith("xls") || tmpFile.getName().endsWith("xlsx")){
                    Converter.convert(inputPath + tmpFile.getName(), outputPath, -1);
                }else if(tmpFile.getName().endsWith("csv")){
                    String pattern = "(.+)_(.+).csv";
                    Pattern r = Pattern.compile(pattern);
                    Matcher m = r.matcher(tmpFile.getName());
                    String fileName;
                    String sheetName;
                    if(m.find()){
                        fileName = m.group(1);
                        sheetName = m.group(2);
                    }
                    else{
                        fileName = Converter.getFileNameNoEx(tmpFile.getName());
                        sheetName = "Sheet1";
                    }
                    Converter.convert(inputPath + tmpFile.getName(), outputPath + fileName, -1, sheetName);
                }else{
                    System.out.println("Invalid input file:" + tmpFile.getAbsolutePath());
                }
            }
        }

    }
}
