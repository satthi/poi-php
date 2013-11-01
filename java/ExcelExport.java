import java.io.*;
import java.awt.image.*;
import java.util.*;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.*;
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
    //13:行の高さ
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
    //24:セル色
    //25:セル色(背景色
    //26:塗りつぶしパターン
    //27:フォント色
    //28:フォントサイズ
    //29:フォント
    //30:イタリックフラグ
    //31:太字フラグ
    //32:打ち消し線フラグ
    //33:アンダーラインフラグ
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
    final static Integer EXCEL_CELL_COLOR = 24;
    final static Integer EXCEL_CELL_BACKCOLOR = 25;
    final static Integer EXCEL_CELL_FILL_PATTERN = 26;
    final static Integer EXCEL_FONT_COLOR = 27;
    final static Integer EXCEL_FONT_SIZE = 28;
    final static Integer EXCEL_FONT = 29;
    final static Integer EXCEL_FONT_ITALIC = 30;
    final static Integer EXCEL_FONT_BOLD = 31;
    final static Integer EXCEL_FONT_STRIKEOUT = 32;
    final static Integer EXCEL_FONT_UNDERLUINE = 33;
    final static Integer EXCEL_IMAGE = 34;
    final static Integer EXCEL_IMAGE_MARGIN_X = 35;
    final static Integer EXCEL_IMAGE_MARGIN_Y = 36;
    final static Integer EXCEL_IMAGE_ENDROW = 37;
    final static Integer EXCEL_IMAGE_ENDCOL = 38;
    final static Integer EXCEL_IMAGE_MARGIN_RX = 39;
    final static Integer EXCEL_IMAGE_MARGIN_RY = 40;
    
    public static void main(String[] args){
        FileInputStream in = null;
        Workbook wb = null;
        if (args.length < 3){
            System.out.println("args none");
            return;
        }

        //シートの読み込み
        try{
            if (!args[0].equals("new_file")){
                in = new FileInputStream(args[0]);
                wb = WorkbookFactory.create(in);
            } else {
                in = null;
                //拡張子の取得
                String ext = getSuffix(args[1]);
                if (ext.equals("xls")){
                    wb = new HSSFWorkbook();
                } else if(ext.equals("xlsx")){
                    wb = new XSSFWorkbook();
                } else {
                    wb = new HSSFWorkbook();
                }
                //このときシートがないので新規シートを一つ作っておく
                wb.createSheet();
            }
        }catch(IOException e){
            System.out.println(e.toString());
        }catch(InvalidFormatException e){
            System.out.println(e.toString());
        }finally{
            try{
                if (in != null){
                    in.close();
                }
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
                        
                        //罫線の色の指定など調整中
                        CellStyle old_style = cell.getCellStyle();
                        CellStyle style = wb.createCellStyle();
                        style.cloneStyleFrom(old_style);
                        
                        if (!stringArray[EXCEL_TOPB_STYLE].equals("")){
                            style.setBorderTop(border_type(stringArray[EXCEL_TOPB_STYLE]));
                        }
                        if (!stringArray[EXCEL_LEFTB_STYLE].equals("")){
                            style.setBorderLeft(border_type(stringArray[EXCEL_LEFTB_STYLE]));
                        }
                        if (!stringArray[EXCEL_RIGHTB_STYLE].equals("")){
                            style.setBorderRight(border_type(stringArray[EXCEL_RIGHTB_STYLE]));
                        }
                        if (!stringArray[EXCEL_BOTTOMB_STYLE].equals("")){
                            style.setBorderBottom(border_type(stringArray[EXCEL_BOTTOMB_STYLE]));
                        }
                        if (!stringArray[EXCEL_TOPB_COLOR].equals("")){
                        	// このブックのカスタムパレットを作成します。
                            style.setTopBorderColor(color_type(stringArray[EXCEL_TOPB_STYLE]));
                        }
                        if (!stringArray[EXCEL_LEFTB_COLOR].equals("")){
                            style.setLeftBorderColor(color_type(stringArray[EXCEL_LEFTB_COLOR]));
                            //style.setBottomBorderColor(IndexedColors.SKY_BLUE.getIndex());
                        }
                        if (!stringArray[EXCEL_RIGHTB_COLOR].equals("")){
                            style.setRightBorderColor(color_type(stringArray[EXCEL_RIGHTB_COLOR]));
                            //style.setLeftBorderColor(IndexedColors.ORANGE.getIndex());
                        }
                        if (!stringArray[EXCEL_BOTTOMB_COLOR].equals("")){
                            style.setBottomBorderColor(color_type(stringArray[EXCEL_BOTTOMB_COLOR]));
                            //style.setRightBorderColor(IndexedColors.BLUE_GREY.getIndex());
                        }
                        cell.setCellStyle(style);
                    }else if(stringArray[EXCEL_TYPE].equals("cell_color")) {
                        CellStyle old_style = cell.getCellStyle();
                        CellStyle style = wb.createCellStyle();
                        style.cloneStyleFrom(old_style);
                        style.setFillForegroundColor(color_type(stringArray[EXCEL_CELL_COLOR]));
                        if(!stringArray[EXCEL_CELL_BACKCOLOR].equals("")) {
                            style.setFillBackgroundColor(color_type(stringArray[EXCEL_CELL_BACKCOLOR]));
                        }
                        style.setFillPattern(cell_fillpattern(stringArray[EXCEL_CELL_FILL_PATTERN]));
                        cell.setCellStyle(style);
                    }else if(stringArray[EXCEL_TYPE].equals("font_setting")) {
                        CellStyle old_style = cell.getCellStyle();
                        CellStyle style = wb.createCellStyle();
                        style.cloneStyleFrom(old_style);
                        
                        Font old_font = wb.getFontAt(old_style.getFontIndex());
                        Font font = wb.createFont();
                        
                        if (!stringArray[EXCEL_FONT_COLOR].equals("")){
                            font.setColor(color_type(stringArray[EXCEL_FONT_COLOR]));
                        } else {
                            font.setColor(old_font.getColor());
                        }
                        if (!stringArray[EXCEL_FONT_SIZE].equals("")){
                            font.setFontHeightInPoints(Short.parseShort(stringArray[EXCEL_FONT_SIZE]));
                        } else {
                            font.setFontHeightInPoints(old_font.getFontHeightInPoints());
                        }
                        if (!stringArray[EXCEL_FONT].equals("")){
                            font.setFontName(stringArray[EXCEL_FONT]);
                        } else {
                            font.setFontName(old_font.getFontName());
                        }
                        if (!stringArray[EXCEL_FONT_ITALIC].equals("")){
                            if (stringArray[EXCEL_FONT_ITALIC].equals("1")){
                                font.setItalic(true);
                            } else if(stringArray[EXCEL_FONT_ITALIC].equals("0")){
                                font.setItalic(false);
                            }
                        } else {
                            font.setItalic(old_font.getItalic());
                        }
                        if (!stringArray[EXCEL_FONT_BOLD].equals("")){
                            if (stringArray[EXCEL_FONT_BOLD].equals("1")){
                                font.setBoldweight(Font.BOLDWEIGHT_BOLD);
                            } else if(stringArray[EXCEL_FONT_BOLD].equals("0")){
                                font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
                            }
                        } else {
                            font.setBoldweight(old_font.getBoldweight());
                        }
                        if (!stringArray[EXCEL_FONT_STRIKEOUT].equals("")){
                            if (stringArray[EXCEL_FONT_STRIKEOUT].equals("1")){
                                font.setStrikeout(true);
                            } else if(stringArray[EXCEL_FONT_STRIKEOUT].equals("0")){
                                font.setStrikeout(false);
                            }
                        } else {
                            font.setStrikeout(old_font.getStrikeout());
                        }
                        if (!stringArray[EXCEL_FONT_UNDERLUINE].equals("")){
                             font.setUnderline(underline_type(stringArray[EXCEL_FONT_UNDERLUINE]));
                        } else {
                            font.setUnderline(old_font.getUnderline());
                        }
                        style.setFont(font);
                        cell.setCellStyle(style);
                    }else if(stringArray[EXCEL_TYPE].equals("add_image")) {
                        File file = new File(stringArray[EXCEL_IMAGE]);
                        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                        
                        BufferedImage img = ImageIO.read(file);
                        ImageIO.write(img, "png", byteArrayOut);
                        
                        Drawing drawing = sheet.createDrawingPatriarch();
                        CreationHelper helper = wb.getCreationHelper();
                        ClientAnchor anchor = helper.createClientAnchor();

                        //画像をセットするセルの指定
                        anchor.setRow1(Integer.parseInt(stringArray[EXCEL_ROW]));
                        anchor.setCol1(Integer.parseInt(stringArray[EXCEL_COL]));
                        //終端のセルを指定している場合。(この場合のみmarginがきく
                        if(!stringArray[EXCEL_IMAGE_ENDROW].equals("") && !stringArray[EXCEL_IMAGE_ENDCOL].equals("")){
                            anchor.setDx1(Integer.parseInt(stringArray[EXCEL_IMAGE_MARGIN_X]));
                            anchor.setDy1(Integer.parseInt(stringArray[EXCEL_IMAGE_MARGIN_Y]));
                            anchor.setRow2(Integer.parseInt(stringArray[EXCEL_IMAGE_ENDROW]));
                            anchor.setDx2(Integer.parseInt(stringArray[EXCEL_IMAGE_MARGIN_RX]));
                            anchor.setCol2(Integer.parseInt(stringArray[EXCEL_IMAGE_ENDCOL]));
                            anchor.setDy2(Integer.parseInt(stringArray[EXCEL_IMAGE_MARGIN_RY]));
                        }
                        int picIndex = wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_PNG);
                        Picture pic = drawing.createPicture(anchor, picIndex);
                        //終端のセルを指定していない場合はresizeをしないと画像が出てこないため
                        if (stringArray[EXCEL_IMAGE_ENDROW].equals("") || stringArray[EXCEL_IMAGE_ENDCOL].equals("")) {
                            pic.resize();
                        }
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

    public static short color_type(String type){
        if (type.equals("AQUA")){
            return IndexedColors.AQUA.getIndex();
        } else if (type.equals("AUTOMATIC")){
            return IndexedColors.AUTOMATIC.getIndex();
        } else if (type.equals("BLACK")){
            return IndexedColors.BLACK.getIndex();
        } else if (type.equals("BLUE")){
            return IndexedColors.BLUE.getIndex();
        } else if (type.equals("BLUE_GREY")){
            return IndexedColors.BLUE_GREY.getIndex();
        } else if (type.equals("BRIGHT_GREEN")){
            return IndexedColors.BRIGHT_GREEN.getIndex();
        } else if (type.equals("BROWN")){
            return IndexedColors.BROWN.getIndex();
        } else if (type.equals("CORAL")){
            return IndexedColors.CORAL.getIndex();
        } else if (type.equals("CORNFLOWER_BLUE")){
            return IndexedColors.CORNFLOWER_BLUE.getIndex();
        } else if (type.equals("DARK_BLUE")){
            return IndexedColors.DARK_BLUE.getIndex();
        } else if (type.equals("DARK_GREEN")){
            return IndexedColors.DARK_GREEN.getIndex();
        } else if (type.equals("DARK_RED")){
            return IndexedColors.DARK_RED.getIndex();
        } else if (type.equals("DARK_TEAL")){
            return IndexedColors.DARK_TEAL.getIndex();
        } else if (type.equals("DARK_YELLOW")){
            return IndexedColors.DARK_YELLOW.getIndex();
        } else if (type.equals("GOLD")){
            return IndexedColors.GOLD.getIndex();
        } else if (type.equals("GREEN")){
            return IndexedColors.GREEN.getIndex();
        } else if (type.equals("GREY_25_PERCENT")){
            return IndexedColors.GREY_25_PERCENT.getIndex();
        } else if (type.equals("GREY_40_PERCENT")){
            return IndexedColors.GREY_40_PERCENT.getIndex();
        } else if (type.equals("GREY_50_PERCENT")){
            return IndexedColors.GREY_50_PERCENT.getIndex();
        } else if (type.equals("GREY_80_PERCENT")){
            return IndexedColors.GREY_80_PERCENT.getIndex();
        } else if (type.equals("LAVENDER")){
            return IndexedColors.LAVENDER.getIndex();
        } else if (type.equals("LEMON_CHIFFON")){
            return IndexedColors.LEMON_CHIFFON.getIndex();
        } else if (type.equals("LIGHT_CORNFLOWER_BLUE")){
            return IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex();
        } else if (type.equals("LIGHT_GREEN")){
            return IndexedColors.LIGHT_GREEN.getIndex();
        } else if (type.equals("LIGHT_ORANGE")){
            return IndexedColors.LIGHT_ORANGE.getIndex();
        } else if (type.equals("LIGHT_TURQUOISE")){
            return IndexedColors.LIGHT_TURQUOISE.getIndex();
        } else if (type.equals("LIGHT_YELLOW")){
            return IndexedColors.LIGHT_YELLOW.getIndex();
        } else if (type.equals("LIME")){
            return IndexedColors.LIME.getIndex();
        } else if (type.equals("MAROON")){
            return IndexedColors.MAROON.getIndex();
        } else if (type.equals("OLIVE_GREEN")){
            return IndexedColors.OLIVE_GREEN.getIndex();
        } else if (type.equals("ORANGE")){
            return IndexedColors.ORANGE.getIndex();
        } else if (type.equals("ORCHID")){
            return IndexedColors.ORCHID.getIndex();
        } else if (type.equals("PALE_BLUE")){
            return IndexedColors.PALE_BLUE.getIndex();
        } else if (type.equals("PINK")){
            return IndexedColors.PINK.getIndex();
        } else if (type.equals("PLUM")){
            return IndexedColors.PLUM.getIndex();
        } else if (type.equals("RED")){
            return IndexedColors.RED.getIndex();
        } else if (type.equals("ROSE")){
            return IndexedColors.ROSE.getIndex();
        } else if (type.equals("ROYAL_BLUE")){
            return IndexedColors.ROYAL_BLUE.getIndex();
        } else if (type.equals("SEA_GREEN")){
            return IndexedColors.SEA_GREEN.getIndex();
        } else if (type.equals("SKY_BLUE")){
            return IndexedColors.SKY_BLUE.getIndex();
        } else if (type.equals("TAN")){
            return IndexedColors.TAN.getIndex();
        } else if (type.equals("TEAL")){
            return IndexedColors.TEAL.getIndex();
        } else if (type.equals("TURQUOISE")){
            return IndexedColors.TURQUOISE.getIndex();
        } else if (type.equals("VIOLET")){
            return IndexedColors.VIOLET.getIndex();
        } else if (type.equals("WHITE")){
            return IndexedColors.WHITE.getIndex();
        } else if (type.equals("YELLOW")){
            return IndexedColors.YELLOW.getIndex();
        } else if (type.equals("AUTOMATIC")){
            return IndexedColors.AUTOMATIC.getIndex();
        } else {
            return IndexedColors.AUTOMATIC.getIndex();
        }
    }
    
    public static byte underline_type(String type){
        if (type.equals("NONE")){
            return Font.U_NONE;
        } else if (type.equals("SINGLE")){
            return Font.U_SINGLE;
        } else if (type.equals("DOUBLE")){
            return Font.U_DOUBLE;
        } else if (type.equals("SINGLE_ACCOUNTING")){
            return Font.U_SINGLE_ACCOUNTING;
        } else if (type.equals("DOUBLE_ACCOUNTING")){
            return Font.U_DOUBLE_ACCOUNTING;
        } else if (type.equals("1")){
            return Font.U_SINGLE;
        } else if (type.equals("0")){
            return Font.U_NONE;
        } else {
            return Font.U_NONE;
        }
    }
    
    public static byte cell_fillpattern(String type){
        if (type.equals("NO_FILL")){
            return CellStyle.NO_FILL;
        } else if (type.equals("SOLID_FOREGROUND")){
            return CellStyle.SOLID_FOREGROUND;
        } else if (type.equals("FINE_DOTS")){
            return CellStyle.FINE_DOTS;
        } else if (type.equals("ALT_BARS")){
            return CellStyle.ALT_BARS;
        } else if (type.equals("SPARSE_DOTS")){
            return CellStyle.SPARSE_DOTS;
        } else if (type.equals("THICK_HORZ_BANDS")){
            return CellStyle.THICK_HORZ_BANDS;
        } else if (type.equals("THICK_VERT_BANDS")){
            return CellStyle.THICK_VERT_BANDS;
        } else if (type.equals("THICK_BACKWARD_DIAG")){
            return CellStyle.THICK_BACKWARD_DIAG;
        } else if (type.equals("THICK_FORWARD_DIAG")){
            return CellStyle.THICK_FORWARD_DIAG;
        } else if (type.equals("BIG_SPOTS")){
            return CellStyle.BIG_SPOTS;
        } else if (type.equals("BRICKS")){
            return CellStyle.BRICKS;
        } else if (type.equals("THIN_HORZ_BANDS")){
            return CellStyle.THIN_HORZ_BANDS;
        } else if (type.equals("THIN_VERT_BANDS")){
            return CellStyle.THIN_VERT_BANDS;
        } else if (type.equals("THIN_BACKWARD_DIAG")){
            return CellStyle.THIN_BACKWARD_DIAG;
        } else if (type.equals("THIN_FORWARD_DIAG")){
            return CellStyle.THIN_FORWARD_DIAG;
        } else if (type.equals("SQUARES")){
            return CellStyle.SQUARES;
        } else if (type.equals("DIAMONDS")){
            return CellStyle.DIAMONDS;
        } else {
            return CellStyle.SOLID_FOREGROUND;
        }
    }
    
    public static String getSuffix(String fileName) {
        if (fileName == null)
            return null;
        int point = fileName.lastIndexOf(".");
        if (point != -1) {
            return fileName.substring(point + 1);
        }
        return fileName;
    }
}
