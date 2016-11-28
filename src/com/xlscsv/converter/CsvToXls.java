package com.xlscsv.converter;
import com.opencsv.CSVReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
class CsvToXls {
    CsvToXls(){

    }

    void process(String inputFileName, String outputFilePath) throws IOException{
        int i=0;
        String s[];
        //ArrayList list= new ArrayList<String>();
        Workbook wb= new HSSFWorkbook();
        Sheet sh= wb.createSheet("Sheet1");
        Row row;
        Cell cell;

        try{
            CSVReader reader = new CSVReader(new FileReader(inputFileName));
            while((s = reader.readNext()) !=null){
                row= sh.createRow(i);
                for(int j=0;j<s.length;j++){
                    cell= row.createCell(j);
                    cell.setCellValue(s[j]);
                }
                i+=1;
            }
        }
        catch(FileNotFoundException e){
            System.out.println("FileNotFound!");}
        catch (IOException e){
            e.printStackTrace();
        }

        String outputFileName = outputFilePath + ".xls";
        FileOutputStream fout= new FileOutputStream(outputFileName);
        wb.write(fout);
        fout.close();
//        System.out.println("Ok!");
    }
}
