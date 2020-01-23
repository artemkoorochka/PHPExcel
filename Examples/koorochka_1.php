<?

$arParams = array();
$arResult = array();

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");

/** Include PHPExcel */
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/../Classes/PHPExcel/IOFactory.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$arResult[] = date('H:i:s') . " Load from Excel2007 file";
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load("data/lansichina/test3.xlsx");

$arResult[] = date('H:i:s') . " Iterate worksheets by Row";

foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
    $arResult[] = 'Worksheet - ' . $worksheet->getTitle();

    foreach ($worksheet->getRowIterator() as $row) {
        $arResult[] = '    Row number - ' . $row->getRowIndex();

        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false); // Loop all cells, even if it is not set
        foreach ($cellIterator as $cell) {
            if (!is_null($cell)) {
                $arResult[] =  '        Cell - ' . $cell->getCoordinate() . ' - ' . $cell->getValue();
            }
        }
    }
}

d($arParams);
d($arResult);
?>