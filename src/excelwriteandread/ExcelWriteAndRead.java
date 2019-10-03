/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelwriteandread;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author hdurmaz
 */
public class ExcelWriteAndRead {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        
     
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet");
      
        
        
        Map<String, Object[]> data = new HashMap<String, Object[]>(); //verileri saklamak için

        data.put("1", new Object[] {"No", "Name", "Salary"});//put: anahtar değer ikilisimi kaydeder.

        data.put("2", new Object[] {1d, "Ali", 5000d});

        data.put("3", new Object[] {2d, "Ayşe", 8000d});

        data.put("4", new Object[] {3d, "Mert", 4000d});
        
        Set<String> keyset = data.keySet(); //verileri bir dizin kullanmadan saklar.

        int rownum = 0;

        for (String key : keyset) {
             Row row = sheet.createRow(rownum++);

             Object [] objArr = data.get(key);

             int cellnum = 0;

             for (Object obj : objArr) {
                 Cell cell = row.createCell(cellnum++);

                 if(obj instanceof Date) 

                 cell.setCellValue((Date)obj);

                 else if(obj instanceof Boolean)

                 cell.setCellValue((Boolean)obj);

                 else if(obj instanceof String)

                 cell.setCellValue((String)obj);

                 else if(obj instanceof Double)

                 cell.setCellValue((Double)obj);
    }

}
        try {
            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\hdurmaz\\test2.xls")); //byte tipinde değişken yazar
             workbook.write(out);
            out.close();
            System.out.println("Excel yazıldı..");
          }


 catch (FileNotFoundException e) {

    e.printStackTrace();

} catch (IOException e) {

    e.printStackTrace();

}
              
        
    }
}

    