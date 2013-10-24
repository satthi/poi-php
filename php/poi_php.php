<?php

class PoiPHP {

    private $export_param = array();
    private $_settings = array();

    /**
     * __construct
     */
    public function __construct() {
    }
    
    public function settings($settings){
        $default_settings = array(
            'poi_path' => dirname(__FILE__) . '/../../poi-3.9',
            'opencsv_path' => dirname(__FILE__) . '/../../opencsv-2.3/deploy/opencsv-2.3.jar',
            'plugin_java_path' => dirname(__FILE__) . '/../java',
        );
        $this->_settings = array_merge($default_settings,$settings);
        //必要ファイルやディレクトリがない場合はエラー
        if (!is_dir($this->_settings['poi_path']) || !file_exists($this->_settings['opencsv_path']) || !is_dir($this->_settings['plugin_java_path'])){
            trigger_error('Java File or Directory Not Found');
            return;
        }
    }


    /**
     * export
     */
    public function export($readfile, $outFile) {
        //$tsv_file = '/tmp/tmp_csv_' . substr((md5(time())), 0, 10) . '.csv';
        $tsv_file = TMP  . 'tmp_csv_' . substr((md5(time())), 0, 10) . '.csv';
        $this->__makeTsv($tsv_file);
        //作ったTSVを元にExcelを作成する
        $cd_command = $this->_settings['plugin_java_path'];
        $command = 'export LANG=ja_JP.UTF-8;cd ' . $cd_command . ';java -Dfile.encoding=UTF-8 -cp \'.:' . $this->_settings['poi_path'] . '/*:' . $this->_settings['poi_path'] . '/lib/*:' . $this->_settings['poi_path'] . '/ooxml-lib/*:' . $this->_settings['opencsv_path'] . '\' ExcelExport ' . $readfile . ' ' . $outFile . ' ' . $tsv_file . ' 2>&1';

        exec($command,$javalog);
        //@unlink($tsv_file);
        
        print_a($javalog);
        if (file_exists($outFile)){
            return $outFile;
        } else {
            return $javalog;
        }
    }
    
    /**
     * import
     */
    public function import($readfile, $outFile, $sheet, $rowst, $colst, $colnum,$file_encode = 'UTF-8') {
        $cd_command = $this->_settings['plugin_java_path'];
        $command = 'export LANG=ja_JP.UTF-8;cd ' . $cd_command . ';java -Dfile.encoding=UTF-8 -cp \'.:' . $this->_settings['poi_path'] . '/*:' . $this->_settings['poi_path'] . '/lib/*:' . $this->_settings['poi_path'] . '/ooxml-lib/*:' . $this->_settings['opencsv_path'] . '\' ExcelImport ' . $readfile . ' ' . $outFile . ' ' . $sheet . ' ' . $rowst . ' ' . $colst . ' ' . $colnum . ' ' . $file_encode . ' 2>&1';

        exec($command,$javalog);
        if (file_exists($outFile)){
            return $outFile;
        } else {
            return $javalog;
        }
    }
    
    public function addString($sheet,$row,$col,$string,$orgrow = null,$orgcol = null,$orgsheet = null){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'val' => $string,
            'type' => 'string',
        );
        if ($orgrow !== null && $orgcol !== null){
            if ($orgsheet === null){
                $orgsheet = $sheet;
            }
            $this->export_param[] = array(
                'sheet' => $sheet,
                'row' => $row,
                'col' => $col,
                'orgsheet' => $orgsheet,
                'orgrow' => $orgrow,
                'orgcol' => $orgcol,
                'type' => 'copy_style',
            );
        }
    }
    
    public function addNumber($sheet,$row,$col,$integer){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'val' => $integer,
            'type' => 'integer',
        );
    }
    
    public function addFormula($sheet,$row,$col,$formula){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'val' => $formula,
            'type' => 'formula',
        );
    }
    
    public function setCellMerge($sheet,$rowst,$rowen,$colst,$colen){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'rowst' => $rowst,
            'rowen' => $rowen,
            'colst' => $colst,
            'colen' => $colen,
            'type' => 'cell_merge',
        );
    }
    
    public function addSheet($org_sheet,$count){
        $this->export_param[] = array(
            'sheet' => $org_sheet,
            'count' => $count,
            'type' => 'sheet_copy',
        );
    }
    
    public function rmSheet($sheet){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'type' => 'sheet_delete',
        );
    }
    
    public function setSheetname($sheet,$sheetname){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'sheetname' => $sheetname,
            'type' => 'sheet_rename',
        );
    }
    
    public function copyCell($sheet,$row,$col,$orgsheet,$orgrow,$orgcol){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'orgsheet' => $orgsheet,
            'orgrow' => $orgrow,
            'orgcol' => $orgcol,
            'type' => 'copy_cell',
        );
    }
    
    public function copyStyle($sheet,$row,$col,$orgsheet,$orgrow,$orgcol){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'orgsheet' => $orgsheet,
            'orgrow' => $orgrow,
            'orgcol' => $orgcol,
            'type' => 'copy_style',
        );
    }
    /*
    public function setStyle($sheet,$row,$col){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'type' => 'set_style',
        );
    }
    */
    public function setBorder($sheet,$row,$col,$topbstyle,$topbcolor,$leftbstyle,$leftbcolor,$rightbstyle,$rightbcolor,$bottombstyle,$bottombcolor){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'topbstyle' => $topbstyle,
            'topbcolor' => $topbcolor,
            'leftbstyle' => $leftbstyle,
            'leftbcolor' => $leftbcolor,
            'rightbstyle' => $rightbstyle,
            'rightbcolor' => $rightbcolor,
            'bottombstyle' => $bottombstyle,
            'bottombcolor' => $bottombcolor,
            'type' => 'set_style',
        );
    }
    public function setCellColor($sheet,$row,$col,$cellcolor,$backcolor = null,$fillpattern = 'SOLID_FOREGROUND'){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'cellcolor' => $cellcolor,
            'backcolor' => $backcolor,
            'fillpattern' => $fillpattern,
            'type' => 'cell_color',
        );
    }
    public function setFontSetting($sheet,$row,$col,$fontcolor = null,$fontsize = null,$font = null,$italic = null,$bold = null,$strikeout = null,$underline = null){
        if ($italic === null){
            $italic_disp = '';
        } else {
            $italic_disp = (int)$italic;
        }
        if ($bold === null){
            $bold_disp = '';
        } else {
            $bold_disp = (int)$bold;
        }
        if ($strikeout === null){
            $strikeout_disp = '';
        } else {
            $strikeout_disp = (int)$strikeout;
        }
        if ($underline === null){
            $underline_disp = '';
        } else {
            if (is_bool($underline)){
                $underline_disp = (int)$underline;
            } else {
                $underline_disp = $underline;
            }
        }
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'col' => $col,
            'fontcolor' => $fontcolor,
            'fontsize' => $fontsize,
            'font' => $font,
            'italic' => $italic_disp,
            'bold' => $bold_disp,
            'strikeout' => $strikeout_disp,
            'underline' => $underline_disp,
            'type' => 'font_setting',
        );
    }
    
    public function chgRowHeight($sheet,$row,$rowheight){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'row' => $row,
            'rowheight' => $rowheight,
            'type' => 'row_height',
        );
    }
    
    public function chgColWidth($sheet,$col,$colwidth){
        $this->export_param[] = array(
            'sheet' => $sheet,
            'col' => $col,
            'colwidth' => $colwidth,
            'type' => 'col_width',
        );
    }
    
    public function reset(){
        $this->export_param = array();
    }

    
    private function __makeTsv($tsv_file){
        //TSVファイル作成
        $tsv_text = '';
        foreach ($this->export_param as $param){
            if ($tsv_text !== ''){
                $tsv_text .= "\r\n";
            }
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
            $tsv_text .= $this->__parseCsv($param['type'],",")  . "," .
                         $this->__parseCsv($param['sheet'],",") . "," .
                         $this->__parseCsv(@$param['row'],",") . "," .
                         $this->__parseCsv(@$param['col'],",") . "," .
                         $this->__parseCsv(@$param['val'],",") . "," .
                         $this->__parseCsv(@$param['rowst'],",") . "," .
                         $this->__parseCsv(@$param['rowen'],",") . "," .
                         $this->__parseCsv(@$param['colst'],",") . "," .
                         $this->__parseCsv(@$param['colen'],",") . "," .
                         $this->__parseCsv(@$param['count'],",") . "," .
                         $this->__parseCsv(@$param['orgsheet'],",") . "," .
                         $this->__parseCsv(@$param['orgrow'],",") . "," .
                         $this->__parseCsv(@$param['orgcol'],",") . "," .
                         $this->__parseCsv(@$param['rowheight'],",") . "," .
                         $this->__parseCsv(@$param['colwidth'],",") . "," .
                         $this->__parseCsv(@$param['sheetname'],",") . "," .
                         $this->__parseCsv(@$param['topbstyle'],",") . "," .
                         $this->__parseCsv(@$param['topbcolor'],",") . "," .
                         $this->__parseCsv(@$param['leftbstyle'],",") . "," .
                         $this->__parseCsv(@$param['leftbcolor'],",") . "," .
                         $this->__parseCsv(@$param['rightbstyle'],",") . "," .
                         $this->__parseCsv(@$param['rightbcolor'],",") . "," .
                         $this->__parseCsv(@$param['bottombstyle'],",") . "," .
                         $this->__parseCsv(@$param['bottombcolor'],",") . "," .
                         $this->__parseCsv(@$param['cellcolor'],",") . "," .
                         $this->__parseCsv(@$param['backcolor'],",") . "," .
                         $this->__parseCsv(@$param['fillpattern'],",") . "," .
                         $this->__parseCsv(@$param['fontcolor'],",") . "," .
                         $this->__parseCsv(@$param['fontsize'],",") . "," .
                         $this->__parseCsv(@$param['font'],",") . "," .
                         $this->__parseCsv(@$param['italic'],",") . "," .
                         $this->__parseCsv(@$param['bold'],",") . "," .
                         $this->__parseCsv(@$param['strikeout'],",") . "," .
                         $this->__parseCsv(@$param['underline'],",")
                        ;
        }
        
        touch($tsv_file);
        $fp = fopen($tsv_file , 'w');
        fwrite($fp,$tsv_text,strlen($tsv_text));
        fclose($fp);
    }
    
    /*
     * __parseCsv
     */
    private function __parseCsv($v, $delimiter) {
        if (preg_match('/[' . $delimiter . '"]/', $v)) {
            $v = str_replace('"', '""', $v);
        }
        $v = '"' . $v . '"';
        return $v;
    }
}