<?
/**
 * https://github.com/PHPOffice/PHPExcel/blob/develop/Documentation/markdown/Overview/04-Configuration-Settings.md
 * system/7studio/excel/git/PHPExcel/Examples/24readfilter.php
 */
if (ini_get('mbstring.func_overload') & 2) {
    ini_set("mbstring.func_overload", 0);
}

require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';

class MyReadFilter implements PHPExcel_Reader_IReadFilter
{
    public function readCell($column, $row, $worksheetName = '') {
        // Read title row and rows 20 - 30
        if ($row == 1 || ($row >= 20 && $row <= 30)) {
            return true;
        }

        return false;
    }
}

$arParams = array();
$arResult = array();

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");

/** Include PHPExcel */


// Create new PHPExcel object
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objReader->setReadFilter( new MyReadFilter() );
$objPHPExcel = $objReader->load($_SERVER["DOCUMENT_ROOT"] . "/upload/excel/sample/test1.xls");


/////////
///


