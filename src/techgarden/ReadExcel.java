/**
 *
 */
package techgarden;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.util.Arrays;
import java.util.Iterator;
//import 

/**
 * @author Жанат
 */
public class ReadExcel {

    public Double[][] getData(String excelFilePath) throws IOException {
        FileInputStream fis = new FileInputStream(new File(excelFilePath));

        Workbook workbook = new XSSFWorkbook(fis);

        Sheet firstSheet = workbook.getSheetAt(0);
        int rownum = firstSheet.getLastRowNum();
        int colnum = firstSheet.getRow(0).getLastCellNum();
        Double[][] data = new Double[rownum][colnum];
        for (int i = 0; i < rownum; i++) {
            Row row = firstSheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < colnum; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        try {
                            data[i][j] = cell.getNumericCellValue();
                        } catch (IllegalStateException e) {
                            e.printStackTrace();
                            //
                        }
                    }
                }
            }
        }
        workbook.close();
        fis.close();
        return data;
    }
/*    catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e){
        e.printStackTrace();
    }*/


    public static void main(String args[]) throws IOException {

        String excelFilePath = "C:\\excel.xlsx";
        ReadExcel readExcel = new ReadExcel();
        Double[][] array1 = readExcel.getData(excelFilePath);
        System.out.println(Arrays.deepToString(array1));
    }

}