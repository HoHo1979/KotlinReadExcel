import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

public class Main {
    public static void main(String[] args) throws IOException {

        InputStream in = new FileInputStream("howinestock.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        XSSFSheet sheet=workbook.getSheetAt(0);
        XSSFRow row= sheet.getRow(6);
        XSSFCell cell = row.getCell(2);
        if(cell.getCellType()== CellType.NUMERIC){
            System.out.println(cell.getNumericCellValue());
        }else if(cell.getCellType() == CellType.STRING){
            System.out.println(cell.getStringCellValue());
        }

    }
}
