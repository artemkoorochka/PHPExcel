<?
if (ini_get('mbstring.func_overload') & 2) {
    ini_set("mbstring.func_overload", 0);
}

$arParams = array();
$arResult = array();

//require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';



$fileName = $_SERVER["DOCUMENT_ROOT"] . "/upload/excel/40b/40b446e24e0fa244d1ec3132a7b763b6.xls";

$inputFileType = 'Excel5';
$inputFileName = $fileName;
$sheetnames = array('Data Sheet #1','Data Sheet #3');
/** Create a new Reader of the type defined in $inputFileType **/
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
/** Advise the Reader of which WorkSheets we want to load **/
$objReader->setLoadSheetsOnly($sheetnames);
/**  Load $inputFileName to a PHPExcel Object  **/
$objPHPExcel = $objReader->load($inputFileName);