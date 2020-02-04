<?
if (ini_get('mbstring.func_overload') & 2) {
    ini_set("mbstring.func_overload", 0);
}

$arParams = array();
$arResult = array();

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_discISAM;
$cacheSettings = array(
    'dir' => $_SERVER["DOCUMENT_ROOT"] . "/upload/excel/cache"
);
PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

// Create new PHPExcel object
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load($_SERVER["DOCUMENT_ROOT"] . "/upload/excel/40b/40b446e24e0fa244d1ec3132a7b763b6.xls");


