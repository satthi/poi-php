import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import au.com.bytecode.opencsv.*;

public class ExcelExport{
    //TSVの書式(使わないものもある
    //0:種別※必須
    //1:シート番号※必須？
    //2:行番号 (値挿入時使用)
    //3:列番号番号(値挿入時使用)
    //4:挿入値(値挿入時使用)
    //5:開始行(セル結合時使用)
    //6:終了行(セル結合時使用)
    //7:開始列(セル結合時使用)
    //8:終了列(セル結合時使用)
    //9:カウント(シートコピー数)
    //10:コピー元シート
    //11:コピー元行
    //12:コピー元列
    //13:列の高さ
    //14:列の幅
    //15:シート名
    //16:上罫線
    //17:上罫線色
    //18:左罫線
    //19:左罫線色
    //20:右罫線
    //21:右罫線色
    //22:下罫線
    //23:下罫線色
    final static Integer EXCEL_TYPE = 0;
    final static Integer EXCEL_SHEET_NO = 1;
    final static Integer EXCEL_ROW = 2;
    final static Integer EXCEL_COL = 3;
    final static Integer EXCEL_VALUE = 4;
    final static Integer EXCEL_ROWST = 5;
    final static Integer EXCEL_ROWEN = 6;
    final static Integer EXCEL_COLST = 7;
    final static Integer EXCEL_COLEN = 8;
    final static Integer EXCEL_COUNT = 9;
    final static Integer EXCEL_ORG_SHEET_NO = 10;
    final static Integer EXCEL_ORG_ROW = 11;
    final static Integer EXCEL_ORG_COL = 12;
    final static Integer EXCEL_ROW_HEIGHT = 13;
    final static Integer EXCEL_COL_WIDTH = 14;
    final static Integer EXCEL_SHEET_NAME = 15;
    final static Integer EXCEL_TOPB_STYLE = 16;
    final static Integer EXCEL_TOPB_COLOR= 17;
    final static Integer EXCEL_LEFTB_STYLE = 18;
    final static Integer EXCEL_LEFTB_COLOR= 19;
    final static Integer EXCEL_RIGHTB_STYLE = 20;
    final static Integer EXCEL_RIGHTB_COLOR = 21;
    final static Integer EXCEL_BOTTOMB_STYLE = 22;
    final static Integer EXCEL_BOTTOMB_COLOR = 23;
    
    public static void main(String[] args){
        FileInputStream in = null;
        Workbook wb = null;
        if (args.length < 3){
            System.out.println("args none");
            return;
        }

        //シートの読み込み
        try{
            in = new FileInputStream(args[0]);
            wb = WorkbookFactory.create(in);
        }catch(IOException e){
            System.out.println(e.toString());
        }catch(InvalidFormatException e){
            System.out.println(e.toString());
        }finally{
            try{
                in.close();
            }catch (IOException e){
                System.out.println(e.toString());
            }
        }

        try {
            CSVReader reader = new CSVReader(new FileReader(args[2]));
            String [] stringArray;
            
            while ((stringArray = reader.readNext()) != null) {
                Sheet sheet = wb.getSheetAt(Integer.parseInt(stringArray[EXCEL_SHEET_NO]));
                //シートのコピー時
                if (stringArray[EXCEL_TYPE].equals("sheet_copy")){
                    for (int i = 0; i < Integer.parseInt(stringArray[EXCEL_COUNT]); i++) {
                        wb.cloneSheet(Integer.parseInt(stringArray[EXCEL_SHEET_NO]));
                    }
                //シート名の変更
                }else if (stringArray[EXCEL_TYPE].equals("sheet_rename")){
                	System.out.println((Integer.parseInt(stringArray[EXCEL_SHEET_NO])));
                	System.out.println((stringArray[EXCEL_SHEET_NAME]));
                    wb.setSheetName(Integer.parseInt(stringArray[EXCEL_SHEET_NO]), stringArray[EXCEL_SHEET_NAME]);
                //シート削除
                }else if (stringArray[EXCEL_TYPE].equals("sheet_delete")){
                    wb.removeSheetAt(Integer.parseInt(stringArray[EXCEL_SHEET_NO]));
                //セルのマージ
                }else if (stringArray[EXCEL_TYPE].equals("cell_merge")){
                    sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(stringArray[EXCEL_ROWST]), Integer.parseInt(stringArray[EXCEL_ROWEN]), Integer.parseInt(stringArray[EXCEL_COLST]), Integer.parseInt(stringArray[EXCEL_COLEN])));
                //行の高さ指定
                }else if (stringArray[EXCEL_TYPE].equals("row_height")){
                    Row row = sheet.getRow(Integer.parseInt(stringArray[EXCEL_ROW]));
                    if (row == null){
                        row = sheet.createRow(Integer.parseInt(stringArray[EXCEL_ROW]));
                    }
                    row.setHeightInPoints(Float.parseFloat(stringArray[EXCEL_ROW_HEIGHT]));
                //列の幅指定
                }else if (stringArray[EXCEL_TYPE].equals("col_width")){
                    /*
                    //エラーが頻発して現在使い物にならないのでいったん削除
                    //エラー回避のために1行目の該当列にデータがなければ空データを作る
                    Row row = sheet.getRow(0);
                    if (row == null){
                        row = sheet.createRow(0);
                    }
                    Cell cell = row.getCell(Integer.parseInt(stringArray[EXCEL_COL]));
                    if (cell == null){
                        cell = row.createCell(Integer.parseInt(stringArray[EXCEL_COL]));
                        cell.setCellValue("dummy");
                    }
                    sheet.setColumnWidth(Integer.parseInt(stringArray[EXCEL_COL]), Integer.parseInt(stringArray[EXCEL_COL_WIDTH]));
                    */
                //セルに値をセットするもの
                } else {
                    Row row = sheet.getRow(Integer.parseInt(stringArray[EXCEL_ROW]));
                    if (row == null){
                        row = sheet.createRow(Integer.parseInt(stringArray[EXCEL_ROW]));
                    }
                    Cell cell = row.getCell(Integer.parseInt(stringArray[EXCEL_COL]));
                    if (cell == null){
                        cell = row.createCell(Integer.parseInt(stringArray[EXCEL_COL]));
                    }
                    
                    if(stringArray[EXCEL_TYPE].equals("copy_cell")) {
                        Sheet org_sheet = wb.getSheetAt(Integer.parseInt(stringArray[EXCEL_ORG_SHEET_NO]));
                        Row org_row = org_sheet.getRow(Integer.parseInt(stringArray[EXCEL_ORG_ROW]));
                        if (org_row == null){
                            org_row = org_sheet.createRow(Integer.parseInt(stringArray[EXCEL_ORG_ROW]));
                        }
                        Cell org_cell = org_row.getCell(Integer.parseInt(stringArray[EXCEL_ORG_COL]));
                        if (org_cell == null){
                            org_cell = org_row.createCell(Integer.parseInt(stringArray[EXCEL_ORG_COL]));
                        }
                        //値をコピーするために、まずセットされている値の型を取得
                        
                        Integer org_cell_type = org_cell.getCellType();
                        if (org_cell_type == Cell.CELL_TYPE_NUMERIC){
                            if (DateUtil.isCellDateFormatted(org_cell)) {
                                cell.setCellValue(org_cell.getDateCellValue());
                            } else {
                                cell.setCellValue(org_cell.getNumericCellValue());
                            }
                            
                        } else if (org_cell_type == Cell.CELL_TYPE_STRING){
                            cell.setCellValue(org_cell.getStringCellValue());
                        } else if (org_cell_type == Cell.CELL_TYPE_FORMULA){
                            cell.setCellValue(org_cell.getCellFormula());
                        } else if (org_cell_type == Cell.CELL_TYPE_BLANK){
                            cell.removeCellComment();
                        } else if (org_cell_type == Cell.CELL_TYPE_BOOLEAN){
                            cell.setCellValue(org_cell.getBooleanCellValue());
                        } else if (org_cell_type == Cell.CELL_TYPE_ERROR){
                            cell.setCellValue(org_cell.getErrorCellValue());
                        }
                    }else if(stringArray[EXCEL_TYPE].equals("copy_style")) {
                        Sheet org_sheet = wb.getSheetAt(Integer.parseInt(stringArray[EXCEL_ORG_SHEET_NO]));
                        Row org_row = org_sheet.getRow(Integer.parseInt(stringArray[EXCEL_ORG_ROW]));
                        if (org_row == null){
                            org_row = org_sheet.createRow(Integer.parseInt(stringArray[EXCEL_ORG_ROW]));
                        }
                        Cell org_cell = org_row.getCell(Integer.parseInt(stringArray[EXCEL_ORG_COL]));
                        if (org_cell == null){
                            org_cell = org_row.createCell(Integer.parseInt(stringArray[EXCEL_ORG_COL]));
                        }
                        //スタイルのコピー
                        cell.setCellStyle(org_cell.getCellStyle());
                    }else if(stringArray[EXCEL_TYPE].equals("set_style")) {
                        /*
                        //罫線の色の指定など調整中
                        CellStyle old_style = cell.getCellStyle();
                        CellStyle style = wb.createCellStyle();
                        style.cloneStyleFrom(old_style);
                        System.out.println(style);
                        
                        if (!stringArray[EXCEL_TOPB_STYLE].equals("")){
                            style.setBorderTop(border_type(stringArray[EXCEL_TOPB_STYLE]));
                        }
                        if (!stringArray[EXCEL_LEFTB_STYLE].equals("")){
                            style.setBorderBottom(CellStyle.BORDER_DOUBLE);
                        }
                        if (!stringArray[EXCEL_RIGHTB_STYLE].equals("")){
                            style.setBorderLeft(CellStyle.BORDER_MEDIUM_DASH_DOT);
                        }
                        if (!stringArray[EXCEL_BOTTOMB_STYLE].equals("")){
                            style.setBorderRight(CellStyle.BORDER_MEDIUM);
                        }
                        if (!stringArray[EXCEL_TOPB_COLOR].equals("")){
                            style.setTopBorderColor(IndexedColors.SKY_BLUE.getIndex());
                        }
                        if (!stringArray[EXCEL_LEFTB_COLOR].equals("")){
                            style.setBottomBorderColor(IndexedColors.SKY_BLUE.getIndex());
                        }
                        if (!stringArray[EXCEL_RIGHTB_COLOR].equals("")){
                            style.setLeftBorderColor(IndexedColors.ORANGE.getIndex());
                        }
                        if (!stringArray[EXCEL_BOTTOMB_COLOR].equals("")){
                            style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
                        }
                        cell.setCellStyle(style);
                        */
                    }else if (stringArray[EXCEL_VALUE] == null || stringArray[EXCEL_VALUE].length() == 0){
                        cell.removeCellComment();
                    } else{
                        if (stringArray[EXCEL_TYPE].equals("string")){
                            cell.setCellValue(stringArray[EXCEL_VALUE]);
                        } else if(stringArray[EXCEL_TYPE].equals("integer")) {
                            cell.setCellValue(Long.parseLong(stringArray[EXCEL_VALUE]));
                        } else if(stringArray[EXCEL_TYPE].equals("formula")) {
                            cell.setCellFormula(stringArray[EXCEL_VALUE]);
                        }
                    }
                }
            }
            

        } catch (FileNotFoundException e) {
            // Fileオブジェクト生成時の例外捕捉
            e.printStackTrace();
        } catch (IOException e) {
            // BufferedReaderオブジェクトのクローズ時の例外捕捉
            e.printStackTrace();
        }
        
        //シートごとの再計算(後回し
        int sheetCount = wb.getNumberOfSheets();
        for (int i = 0; i < sheetCount;i++){
            Sheet recalc_sheet = wb.getSheetAt(i);
            recalc_sheet.setForceFormulaRecalculation(true);
        }
    
        try{
            FileOutputStream out = new FileOutputStream(args[1]);
            wb.write(out);
        }catch(IOException e){
            System.out.println(e.toString());
        }finally{
        }
        
    }
    
    
    
    
    public static short border_type(String type){
        if (type.equals("none")){
            return CellStyle.BORDER_NONE;
        } else if (type.equals("thin")){
            return CellStyle.BORDER_THIN;
        } else if (type.equals("medium")){
            return CellStyle.BORDER_MEDIUM;
        } else if (type.equals("dashed")){
            return CellStyle.BORDER_DASHED;
        } else if (type.equals("dotted")){
            return CellStyle.BORDER_DOTTED;
        } else if (type.equals("thick")){
            return CellStyle.BORDER_THICK;
        } else if (type.equals("dobble")){
            return CellStyle.BORDER_DOUBLE;
        } else if (type.equals("hair")){
            return CellStyle.BORDER_HAIR;
        } else if (type.equals("medium_dashed")){
            return CellStyle.BORDER_MEDIUM_DASHED;
        } else if (type.equals("dash_dot")){
            return CellStyle.BORDER_DASH_DOT;
        } else if (type.equals("medium_dash_dot")){
            return CellStyle.BORDER_MEDIUM_DASH_DOT;
        } else if (type.equals("dash_dot_dot")){
            return CellStyle.BORDER_DASH_DOT_DOT;
        } else if (type.equals("medium_dash_dot_dot")){
            return CellStyle.BORDER_MEDIUM_DASH_DOT_DOT;
        } else if (type.equals("slanted_dash_dot")){
            return CellStyle.BORDER_SLANTED_DASH_DOT;
        } else {
            return CellStyle.BORDER_NONE;
        }
    }
}
