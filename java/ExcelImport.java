import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.*;
import au.com.bytecode.opencsv.*;

public class ExcelImport{
    static public final String DATE_PATTERN ="yyyy-MM-dd'T'HH:mm:ss";
    
    public static void main(String[] args){
        //シートの読み込み
        try{
            FileInputStream in = null;
            Workbook wb = null;
            if (args.length < 7){
                System.out.println("args none");
                return;
            }
            in = new FileInputStream(args[0]);
            wb = WorkbookFactory.create(in);
            Sheet sheet = wb.getSheetAt(Integer.parseInt(args[2]));
            //System.out.println(sheet);
            Integer lastRow = sheet.getLastRowNum();
            //System.out.println(lastRow);
            //String[][] stringArray = new String[lastRow][Integer.parseInt(args[5]) + Integer.parseInt(args[4])];
            List<String[]> strList = new ArrayList<String[]>();
            for (Integer i = Integer.parseInt(args[3]);i <= lastRow;i++){
                Row row = sheet.getRow(i);
                String[] stringArray = new String[Integer.parseInt(args[5]) + Integer.parseInt(args[4])];
                for (Integer j = Integer.parseInt(args[4]);j < Integer.parseInt(args[4]) + Integer.parseInt(args[5]);j++){
                    //値をコピーするために、まずセットされている値の型を取得
                    Cell cell = row.getCell(j);
                    if (cell == null){
                        cell = row.createCell(j);
                    }
                    
                    Integer cell_type = cell.getCellType();
                    String cell_value = "";
                    if (cell_type == Cell.CELL_TYPE_NUMERIC){
                        if (DateUtil.isCellDateFormatted(cell)) {
                            cell_value = (new SimpleDateFormat(DATE_PATTERN)).format(cell.getDateCellValue());
                        } else {
                            DecimalFormat format = new DecimalFormat("0.#");
                            cell_value = format.format(cell.getNumericCellValue());
                        }
                        
                    } else if (cell_type == Cell.CELL_TYPE_STRING){
                        cell_value = cell.getStringCellValue();
                    } else if (cell_type == Cell.CELL_TYPE_FORMULA){
                        cell_value = cell.getCellFormula();
                    } else if (cell_type == Cell.CELL_TYPE_BLANK){
                        cell_value = "";
                    } else if (cell_type == Cell.CELL_TYPE_BOOLEAN){
                        cell_value = String.valueOf(cell.getBooleanCellValue());
                    } else if (cell_type == Cell.CELL_TYPE_ERROR){
                        //String cell_value = new String(cell.getErrorCellValue(), "UTF-8");
                        cell_value = "";
                    } else {
                        cell_value = "";
                    }
                    stringArray[j] = cell_value;
                }
                strList.add(stringArray);
            }

            FileOutputStream out = new FileOutputStream(args[1]);
            Writer writer = new OutputStreamWriter(out,args[6]);
            CSVWriter csvWriter = new CSVWriter(writer);
            csvWriter.writeAll(strList);
            
            csvWriter.close();
            writer.close();
            out.close();
        }catch(IOException e){
            System.out.println(e.toString());
        }catch(InvalidFormatException e){
            System.out.println(e.toString());
        }finally{
        }
    }
}
