<?
if (ini_get('mbstring.func_overload') & 2) {
    ini_set("mbstring.func_overload", 0);
}

$arParams = array();
$arResult = array();

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';



$fileName = $_SERVER["DOCUMENT_ROOT"] . "/upload/excel/40b/40b446e24e0fa244d1ec3132a7b763b6.xls";
$inputFileName = $fileName;
$fileContent = "";

//get inputFileType (most of time Excel5)
$inputFileType = PHPExcel_IOFactory::identify($inputFileName);

//initialize cache, so the phpExcel will not throw memory overflow
$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
$cacheSettings = array(' memoryCacheSize ' => '8MB');
PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

//initialize object reader by file type
$objReader = PHPExcel_IOFactory::createReader($inputFileType);

//read only data (without formating) for memory and time performance
$objReader->setReadDataOnly(true);

//load file into PHPExcel object
$objPHPExcel = $objReader->load($inputFileName);