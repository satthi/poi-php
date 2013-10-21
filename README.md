***********************************************************************
poi-php
varsion 0.1
(2013/10/21)
���̃v���O�C����Java��poi��PHP����@���ăG�N�Z������o�͂���v���O�C���ł��B
***********************************************************************

�K�{�v��
jdk1.7.0_40
poi-3.9
opencsv-2.3

�g����

�@poi-php��C�ӂ̏ꏊ�ɐݒu
�APHP���L�q
<pre>
//�f�t�H���g�ǂݍ���
require_once('poi-php.php�̃f�B���N�g���p�X');
$this->PoiPHP = new PoiPHP();
$this->PoiPHP->settings(array(
	'poi_path' => 'poi-3.9�̃f�B���N�g���p�X',
	'opencsv_path' => 'opencsv-2.3.jar�̃t�@�C���p�X',
));
//Excel�o��
//1�V�[�g�ڂ�1�s��1���a�𕶎���Ƃ��ē���
$this->ExcelExport->addString(0,0,0,'a');
//���o�͂̃t�@�C���̓t���p�X�Ŏw�肷��B
$readfile = dirname(__FILE__) . '/test.xls';
$outFile = dirname(__FILE__) . '/export.xls';
$this->ExcelExport->export($readfile, $outFile);

//Excel����
//���o�͂̃t�@�C���̓t���p�X�Ŏw�肷��B
$readfile = dirname(__FILE__) . '/test.xls';
$outFile = dirname(__FILE__) . '/export.csv';
$this->ExcelImport->import($readfile, $outFile, 0, 2, 1, 4);
</pre>
���L�q�B

�����ӓ_
Java�̃o�[�W����������Ȃ��Ɠ��삵�܂���B
/java�f�B���N�g�����ɂ�.java�t�@�C�����u���Ă���̂ŁA
java�̃o�[�W���������킹���Ȃ��ꍇ�ɂ͎����ŃR���p�C�������瓮����������Ȃ��ł��B

-----------�ȉ��֐��̐����ł�----------

<pre>
/*
 * addString
 * ������̒ǉ�
 * �Q�ƃZ���̐ݒ������ƃX�^�C�����R�s�[���Ă��܂��B
 * 
 * $sheet �V�[�g�ԍ�
 * $row �s�ԍ�
 * $col ��ԍ�
 * $string ������
 * $orgrow �Q�ƃZ���̍s�ԍ�
 * $orgcol �Q�ƃZ���̗�ԍ�
 * $orgsheet �Q�ƃZ���̃V�[�g�ԍ�
 */
public function addString($sheet,$row,$col,$string,$orgrow = null,$orgcol = null,$orgsheet = null){

/*
 * addNumber
 * ���l�̒ǉ�
 * �Q�ƃZ���̐ݒ������ƃX�^�C�����R�s�[���Ă��܂��B
 * 
 * $sheet �V�[�g�ԍ�
 * $row �s�ԍ�
 * $col ��ԍ�
 * $integer ���l
 * $orgrow �Q�ƃZ���̍s�ԍ�
 * $orgcol �Q�ƃZ���̗�ԍ�
 * $orgsheet �Q�ƃZ���̃V�[�g�ԍ�
 */
public function addNumber($sheet,$row,$col,$integer,$orgrow = null,$orgcol = null,$orgsheet = null)

/*
 * addFormula
 * ���l�̒ǉ�
 * �Q�ƃZ���̐ݒ������ƃX�^�C�����R�s�[���Ă��܂��B
 * 
 * $sheet �V�[�g�ԍ�
 * $row �s�ԍ�
 * $col ��ԍ�
 * $formula �֐�
 * $orgrow �Q�ƃZ���̍s�ԍ�
 * $orgcol �Q�ƃZ���̗�ԍ�
 * $orgsheet �Q�ƃZ���̃V�[�g�ԍ�
 */
public function addFormula($sheet,$row,$col,$formula,$orgrow = null,$orgcol = null,$orgsheet = null)


/*
 * setCellMerge
 * �Z���̃}�[�W
 * 
 * $sheet �V�[�g�ԍ�
 * $rowst �J�n�s�ԍ�
 * $rowen �I���s�ԍ�
 * $colst �J�n��ԍ�
 * $colen �I����ԍ�
 */
public function setCellMerge($sheet,$rowst,$rowen,$colst,$colen)

/*
 * addSheet
 * �V�[�g�̒ǉ�
 * 
 * $org_sheet �匳�̃V�[�g�ԍ�
 * $count �V�[�g�ǉ���
 */
public function addSheet($org_sheet,$count)


/*
 * rmSheet
 * �V�[�g�̍폜
 * 
 * $sheet �V�[�g�ԍ�
 */
public function rmSheet($sheet)

/*
 * setSheetname
 * �V�[�g���̐ݒ�
 * 
 * $sheet �V�[�g�ԍ�
 * $sheet �V�[�g��
 */
public function setSheetname($sheet,$sheetname)

/*
 * copyCell
 * �Z���̃R�s�[
 * 
 * $sheet �V�[�g�ԍ�
 * $row �s�ԍ�
 * $col ��ԍ�
 * $orgrow �Q�ƃZ���̍s�ԍ�
 * $orgcol �Q�ƃZ���̗�ԍ�
 * $orgsheet �Q�ƃZ���̃V�[�g�ԍ�
 */
public function copyCell($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * copyStyle
 * �X�^�C���̃R�s�[
 * 
 * $sheet �V�[�g�ԍ�
 * $row �s�ԍ�
 * $col ��ԍ�
 * $orgrow �Q�ƃZ���̍s�ԍ�
 * $orgcol �Q�ƃZ���̗�ԍ�
 * $orgsheet �Q�ƃZ���̃V�[�g�ԍ�
 */
public function copyStyle($sheet,$row,$col,$orgsheet,$orgrow,$orgcol)

/*
 * reset
 * �Z�b�g�����l�̃��Z�b�g
 */
public function reset()


/*
 * export
 * Excel�̏o�́BaddString�Ȃǂ̐ݒ�����ׂčs���čŌ�ɏo�͂����܂��B
 *
 * @param readfile �ǂݍ��݃e���v���[�g�t�@�C��
 * @param outFile �o��Excel�t�@�C��
 */
public function export($readfile, $outFile)

/*
 * import
 * Excel�̓��́BCSV�t�@�C���ɏo�͂����܂��B
 * �C�ӂ̃V�[�g�̔C�ӂ̏ꏊ����ǂݍ��݂܂��B
 * 
 * $readfile �ǂݍ���Excel�t�@�C��
 * $outFile �o��CSV�t�@�C��
 * $sheet �ǂݍ��ރV�[�g�ԍ�
 * $rowst �ǂݍ��݊J�n�s
 * $colst �ǂݍ��݊J�n��
 * $colnum ��
 * $file_encode CSV�t�@�C���̕����R�[�h
 */
public function import($readfile, $outFile, $sheet, $rowst, $colst, $colnum,$file_encode = 'UTF-8')

</pre>

## License ##

The MIT Lisence

Copyright (c) 2013 Fusic Co., Ltd. (http://fusic.co.jp)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Author ##

Satoru Hagiwara