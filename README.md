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
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$string 文字列
 * (int)$orgrow 参照セルの行番号
 * (int)$orgcol 参照セルの列番号
 * (int)$orgsheet 参照セルのシート番号
 */
public function addString($sheet,$row,$col,$string,$orgrow = null,$orgcol = null,$orgsheet = null){

/*
 * addNumber
 * 数値の追加
 * 参照セルの設定を入れるとスタイルをコピーしてきます。
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (int)$integer 数値
 * (int)$orgrow 参照セルの行番号
 * (int)$orgcol 参照セルの列番号
 * (int)$orgsheet 参照セルのシート番号
 */
public function addNumber($sheet,$row,$col,$integer,$orgrow = null,$orgcol = null,$orgsheet = null)

/*
 * addFormula
 * 数値の追加
 * 参照セルの設定を入れるとスタイルをコピーしてきます。
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$formula 関数
 * (int)$orgrow 参照セルの行番号
 * (int)$orgcol 参照セルの列番号
 * (int)$orgsheet 参照セルのシート番号
 */
public function addFormula($sheet,$row,$col,$formula,$orgrow = null,$orgcol = null,$orgsheet = null)


/*
 * setCellMerge
 * セルのマージ
 * 
 * (int)$sheet シート番号
 * (int)$rowst 開始行番号
 * (int)$rowen 終了行番号
 * (int)$colst 開始列番号
 * (int)$colen 終了列番号
 */
public function setCellMerge($sheet,$rowst,$rowen,$colst,$colen)

/*
 * addSheet
 * シートの追加
 * 
 * (int)$org_sheet 大元のシート番号
 * (int)$count シート追加数
 */
public function addSheet($org_sheet,$count)


/*
 * rmSheet
 * シートの削除
 * 
 * (int)$sheet シート番号
 */
public function rmSheet($sheet)

/*
 * setSheetname
 * シート名の設定
 * 
 * (int)$sheet シート番号
 * (int)$sheet シート名
 */
public function setSheetname($sheet,$sheetname)

/*
 * copyCell
 * セルのコピー
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (int)$orgrow 参照セルの行番号
 * (int)$orgcol 参照セルの列番号
 * (int)$orgsheet 参照セルのシート番号
 */
public function copyCell($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * copyStyle
 * スタイルのコピー
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (int)$orgrow 参照セルの行番号
 * (int)$orgcol 参照セルの列番号
 * (int)$orgsheet 参照セルのシート番号
 */
public function copyStyle($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * setBorder
 * 罫線の設定
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$topbstyle 上罫線の設定
 * (string)$topbcolor 上罫線の色
 * (string)$leftbstyle 左罫線の設定
 * (string)$leftbcolor 左罫線の色
 * (string)$rightbstyle 右罫線の設定
 * (string)$rightbcolor 右罫線の色
 * (string)$bottombstyle 下罫線の設定
 * (string)$bottombcolor 下罫線の色
 * 色の種類、罫線の種類は下記参照
 */
public function setBorder($sheet,$row,$col,$topbstyle,$topbcolor,$leftbstyle,$leftbcolor,$rightbstyle,$rightbcolor,$bottombstyle,$bottombcolor){

/*
 * setCellColor
 * セルの色設定
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$cellcolor セルの色(前景色)の設定
 * (string)$backcolor 背景色の設定
 * (string)$fillpattern 塗りつぶしパターンの設定
 * 色の種類、塗りつぶしパターンの種類は下記参照
 */
public function setCellColor($sheet,$row,$col,$cellcolor,$backcolor = null,$fillpattern = 'SOLID_FOREGROUND')


/*
 * setFontSetting
 * セルの色設定
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$fontcolor フォントの色設定
 * (int)$fontsize フォントのサイズ
 * (string)$font フォントの設定 (MS コジック)など文字列で
 * (bool)$italic イタリックの設定
 * (bool)$bold 太字の設定
 * (bool)$strikeout 打ち消し線の設定
 * (string)$underline 下線の設定
 * 色の種類、下線の種類は下記参照
 */
public function setFontSetting($sheet,$row,$col,$fontcolor = null,$fontsize = null,$font = null,$italic = null,$bold = null,$strikeout = null,$underline = null)


/*
 * addImage
 * 画像の追加
 * 基本は画像パスまで。
 * それ以降のパスはすべて指定をしないと動作しない上、強引に画像が引き伸ばされるため奇麗に画像が表示されないので
 * resizeは別途行った上で設置を推奨
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$image 画像パス
 * (int)$margin_x 左マージン
 * (int)$margin_y 上マージン
 * (int)$endrow 終端行
 * (int)$endcol 終端列
 * (int)$margin_rx 右マージン
 * (int)$margin_ry 下マージン
 */
public function addImage($sheet,$row,$col,$image,$margin_x = 0,$margin_y = 0,$endrow = null,$endcol = null, $margin_rx = 0,$margin_ry = 0){

/*
 * setAlign
 * セルの寄せ設定
 * 
 * (int)$sheet シート番号
 * (int)$row 行番号
 * (int)$col 列番号
 * (string)$align 寄せ
 * 寄せの種類は下記参照
 */
public function setAlign($sheet,$row,$col,$align)

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

色
<pre>
AQUA
AUTOMATIC
BLACK
BLUE
BLUE_GREY
BRIGHT_GREEN
BROWN
CORAL
CORNFLOWER_BLUE
DARK_BLUE
DARK_GREEN
DARK_RED
DARK_TEAL
DARK_YELLOW
GOLD
GREEN
GREY_25_PERCENT
GREY_40_PERCENT
GREY_50_PERCENT
GREY_80_PERCENT
LAVENDER
LEMON_CHIFFON
LIGHT_CORNFLOWER_BLUE
LIGHT_GREEN
LIGHT_ORANGE
LIGHT_TURQUOISE
LIGHT_YELLOW
LIME
MAROON
OLIVE_GREEN
ORANGE
ORCHID
PALE_BLUE
PINK
PLUM
RED
ROSE
ROYAL_BLUE
SEA_GREEN
SKY_BLUE
TAN
TEAL
TURQUOISE
VIOLET
WHITE
YELLOW
AUTOMATIC
</pre>


罫線
<pre>
none
thin
medium
dashed
dotted
thick
dobble
hair
medium_dashed
dash_dot
medium_dash_dot
dash_dot_dot
medium_dash_dot_dot
slanted_dash_dot
</pre>

下線
<pre>
NONE
SINGLE
DOUBLE
SINGLE_ACCOUNTING
DOUBLE_ACCOUNTING
</pre>

塗りつぶし
<pre>
NO_FILL
SOLID_FOREGROUND
FINE_DOTS
ALT_BARS
SPARSE_DOTS
THICK_HORZ_BANDS
THICK_VERT_BANDS
THICK_BACKWARD_DIAG
THICK_FORWARD_DIAG
BIG_SPOTS
BRICKS
THIN_HORZ_BANDS
THIN_VERT_BANDS
THIN_BACKWARD_DIAG
THIN_FORWARD_DIAG
SQUARES
DIAMONDS
</pre>

寄せ
<pre>
LEFT
RIGHT
CENTER
GENERAL
FILL
JUSTIFY
CENTER_SELECTION
</pre>

## License ##

The MIT Lisence

Copyright (c) 2013 Fusic Co., Ltd. (http://fusic.co.jp)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Author ##

Satoru Hagiwara