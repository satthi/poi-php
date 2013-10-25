***********************************************************************
poi-php

varsion 0.1

(2013/10/21)

このプラグインはJavaのpoiをPHPから叩いてエクセルを入出力するプラグインです。
***********************************************************************

必須要件
jdk1.7.0_40
poi-3.9
opencsv-2.3

使い方

①poi-phpを任意の場所に設置
②PHPを記述
<pre>
//デフォルト読み込み
require_once('poi-php.phpのディレクトリパス');
$this->PoiPHP = new PoiPHP();
$this->PoiPHP->settings(array(
	'poi_path' => 'poi-3.9のディレクトリパス',
	'opencsv_path' => 'opencsv-2.3.jarのファイルパス',
));
//Excel出力
//1シート目の1行目1列にaを文字列として入力
$this->ExcelExport->addString(0,0,0,'a');
//入出力のファイルはフルパスで指定する。
$readfile = dirname(__FILE__) . '/test.xls';
$outFile = dirname(__FILE__) . '/export.xls';
$this->ExcelExport->export($readfile, $outFile);

//Excel入力
//入出力のファイルはフルパスで指定する。
$readfile = dirname(__FILE__) . '/test.xls';
$outFile = dirname(__FILE__) . '/export.csv';
$this->ExcelImport->import($readfile, $outFile, 0, 2, 1, 4);
</pre>
を記述。

※注意点
Javaのバージョンが合わないと動作しません。

/javaディレクトリ内には.javaファイルも置いているので、

javaのバージョンを合わせられない場合には自分でコンパイルしたら動くかもしれないです。

-----------以下関数の説明です----------

<pre>
/*
 * addString
 * 文字列の追加
 * 参照セルの設定を入れるとスタイルをコピーしてきます。
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $string 文字列
 * $orgrow 参照セルの行番号
 * $orgcol 参照セルの列番号
 * $orgsheet 参照セルのシート番号
 */
public function addString($sheet,$row,$col,$string,$orgrow = null,$orgcol = null,$orgsheet = null){

/*
 * addNumber
 * 数値の追加
 * 参照セルの設定を入れるとスタイルをコピーしてきます。
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $integer 数値
 * $orgrow 参照セルの行番号
 * $orgcol 参照セルの列番号
 * $orgsheet 参照セルのシート番号
 */
public function addNumber($sheet,$row,$col,$integer,$orgrow = null,$orgcol = null,$orgsheet = null)

/*
 * addFormula
 * 数値の追加
 * 参照セルの設定を入れるとスタイルをコピーしてきます。
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $formula 関数
 * $orgrow 参照セルの行番号
 * $orgcol 参照セルの列番号
 * $orgsheet 参照セルのシート番号
 */
public function addFormula($sheet,$row,$col,$formula,$orgrow = null,$orgcol = null,$orgsheet = null)


/*
 * setCellMerge
 * セルのマージ
 * 
 * $sheet シート番号
 * $rowst 開始行番号
 * $rowen 終了行番号
 * $colst 開始列番号
 * $colen 終了列番号
 */
public function setCellMerge($sheet,$rowst,$rowen,$colst,$colen)

/*
 * addSheet
 * シートの追加
 * 
 * $org_sheet 大元のシート番号
 * $count シート追加数
 */
public function addSheet($org_sheet,$count)


/*
 * rmSheet
 * シートの削除
 * 
 * $sheet シート番号
 */
public function rmSheet($sheet)

/*
 * setSheetname
 * シート名の設定
 * 
 * $sheet シート番号
 * $sheet シート名
 */
public function setSheetname($sheet,$sheetname)

/*
 * copyCell
 * セルのコピー
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $orgrow 参照セルの行番号
 * $orgcol 参照セルの列番号
 * $orgsheet 参照セルのシート番号
 */
public function copyCell($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * copyStyle
 * スタイルのコピー
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $orgrow 参照セルの行番号
 * $orgcol 参照セルの列番号
 * $orgsheet 参照セルのシート番号
 */
public function copyStyle($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * setBorder
 * 罫線の設定
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $topbstyle 上罫線の設定
 * $topbcolor 上罫線の色
 * $leftbstyle 左罫線の設定
 * $leftbcolor 左罫線の色
 * $rightbstyle 右罫線の設定
 * $rightbcolor 右罫線の色
 * $bottombstyle 下罫線の設定
 * $bottombcolor 下罫線の色
 */
public function setBorder($sheet,$row,$col,$topbstyle,$topbcolor,$leftbstyle,$leftbcolor,$rightbstyle,$rightbcolor,$bottombstyle,$bottombcolor){

/*
 * setCellColor
 * セルの色設定
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $cellcolor セルの色(前景色)の設定
 * $backcolor 背景色の設定
 * $fillpattern 塗りつぶしパターンの設定
 */
public function setCellColor($sheet,$row,$col,$cellcolor,$backcolor = null,$fillpattern = 'SOLID_FOREGROUND')


/*
 * setFontSetting
 * セルの色設定
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $fontcolor フォントの色設定
 * $fontsize フォントのサイズ
 * $font フォントの設定
 * $italic イタリックの設定
 * $bold 太字の設定
 * $strikeout 打ち消し線の設定
 * $underline 下線の設定
 */
public function setFontSetting($sheet,$row,$col,$fontcolor = null,$fontsize = null,$font = null,$italic = null,$bold = null,$strikeout = null,$underline = null)


/*
 * addImage
 * 画像の追加
 * 基本は画像パスまで。
 * それ以降のパスはすべて指定をしないと動作しない上、強引に画像が引き伸ばされるため奇麗に画像が表示されないので
 * resizeは別途行った上で設置を推奨
 * 
 * $sheet シート番号
 * $row 行番号
 * $col 列番号
 * $image 画像パス
 * $margin_x 左マージン
 * $margin_y 上マージン
 * $endrow 終端行
 * $endcol 終端列
 * $margin_rx 右マージン
 * $margin_ry 下マージン
 */
public function addImage($sheet,$row,$col,$image,$margin_x = 0,$margin_y = 0,$endrow = null,$endcol = null, $margin_rx = 0,$margin_ry = 0){

/*
 * reset
 * セットした値のリセット
 */
public function reset()


/*
 * excelExport
 * Excelの出力。addStringなどの設定をすべて行って最後に出力をします。
 *
 * @param readfile 読み込みテンプレートファイル
 * @param outFile 出力Excelファイル
 */
public function excelExport($readfile, $outFile)

/*
 * excelImport
 * Excelの入力。CSVファイルに出力をします。
 * 任意のシートの任意の場所から読み込みます。
 * 
 * $readfile 読み込みExcelファイル
 * $outFile 出力CSVファイル
 * $sheet 読み込むシート番号
 * $rowst 読み込み開始行
 * $colst 読み込み開始列
 * $colnum 列数
 * $file_encode CSVファイルの文字コード
 */
public function excelImport($readfile, $outFile, $sheet, $rowst, $colst, $colnum,$file_encode = 'UTF-8')

</pre>

## License ##

The MIT Lisence

Copyright (c) 2013 Fusic Co., Ltd. (http://fusic.co.jp)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Author ##

Satoru Hagiwara