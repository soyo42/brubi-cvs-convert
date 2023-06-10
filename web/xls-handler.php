<?php
if (array_key_exists('debug', $_REQUEST)) {
    $debug_only = boolval($_REQUEST['debug']);
} else {
    $debug_only = false;
}
// printf("debug_only : [%b] <br/>", $debug_only);
// print_r($_REQUEST);
// exit();

// printf("jahoda %s \n", '1');
// $val = 2;
// printf("jahoda $val\n");

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator;
use PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator;
use PhpOffice\PhpSpreadsheet\Shared\Date as XlsDate;

use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;

$sadzby_xls = './sadzby_DPH_s_predkontaciou.xls';
$template_xls = './result_template.xls';
$items_file = './EtsySoldOrderItems2023-3.csv';
$orders_file = './EtsySoldOrders2023-3.csv';

// load sadzby via xls file
$sadzby_spreadsheet = IOFactory::load($sadzby_xls);
// load template xls
$template_spreadsheet = IOFactory::load($template_xls);

if ($debug_only) print("<pre>\n");

// read sadzby
$sadzby_active_worksheet = $sadzby_spreadsheet->getActiveSheet();
$SADZBY = array();
for ($i = 1; true; $i++) {
    $country = $sadzby_active_worksheet->getCell("A$i")->getValue();
    $country_isocode = $sadzby_active_worksheet->getCell("B$i")->getValue();
    if ($country == null) {
        break;
    }
    $c = $sadzby_active_worksheet->getCell("C$i")->getValue();
    $d = $sadzby_active_worksheet->getCell("D$i")->getValue();
    // print("country: $country [$country_isocode] \n");
    $SADZBY[$country_isocode] = array('C' => $c, 'D' => $d, 'country' => $country);
    // $width = 15 - mb_strlen($country) + strlen($country);
    // printf("%-{$width}s :: $c -- $d\n", $country);
}

if ($debug_only) {
    print("--sadzby--\n");
    foreach ($SADZBY as $key => $value) {
        printf("@@ %s [%s | %s] -- %s\n", $key, $value['C'], $value['D'], $value['country']);
    }
}

if ($debug_only) print("----\n");
// load csv file
$reader = new CsvReader();
$reader->setDelimiter(',')
    ->setEnclosure('"')
    ->setSheetIndex(0);

// process orders and items
$orders_speadsheet = $reader->load($orders_file);
$orders_active_worksheet = $orders_speadsheet->getActiveSheet();
$template_active_worksheet = $template_spreadsheet->getActiveSheet();

$colIterator1 = new ColumnCellIterator($orders_active_worksheet, 'A', 2);
foreach ($colIterator1 as $cell) {
    $row = $cell->getRow();
    if ($debug_only) printf("--- %03d %s ---\n", $row, $cell->getValue());

    $target_row = $row + 3;
    $template_active_worksheet->insertNewRowBefore($target_row);
    $template_active_worksheet->getRowDimension($target_row)->setRowHeight(15);


    $cell_date = DateTime::createFromFormat('m/d/y', $cell->getValue());

    if ($debug_only) {
        printf("%s -> %s -- %s\n",
               $template_active_worksheet->getCell('A'.($target_row-1))->getDataType(),
               $cell_date->format('m/d/Y'),
               XlsDate::PHPToExcel($cell_date)
        );
    }
    
    $template_active_worksheet->setCellValue('A'.$target_row, XlsDate::PHPToExcel($cell_date));  // date
    $template_active_worksheet->setCellValue('B'.$target_row, $orders_active_worksheet->getCell('B'.$row)->getValue());  // order no.
    $template_active_worksheet->setCellValue('C'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // full name
    //  item name
    $template_active_worksheet->setCellValue('E'.$target_row, $orders_active_worksheet->getCell('G'.$row)->getValue());  // number of items
    $template_active_worksheet->setCellValue('F'.$target_row, $orders_active_worksheet->getCell('J'.$row)->getValue());  // street name
    $template_active_worksheet->setCellValue('G'.$target_row, $orders_active_worksheet->getCell('L'.$row)->getValue());  // city
    $template_active_worksheet->setCellValue('H'.$target_row, $orders_active_worksheet->getCell('N'.$row)->getValue());  // zip code
    $template_active_worksheet->setCellValue('I'.$target_row, $orders_active_worksheet->getCell('O'.$row)->getValue());  // country
    $template_active_worksheet->setCellValue('J'.$target_row, $orders_active_worksheet->getCell('P'.$row)->getValue());  // currency
    // $template_active_worksheet->setCellValue('K'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // OSS - <druh dodani hlavicka>
    // $template_active_worksheet->setCellValue('L'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // OSS - country
    $template_active_worksheet->setCellValue('M'.$target_row, $orders_active_worksheet->getCell('X'.$row)->getValue());  // OSS - base rate
    // $template_active_worksheet->setCellValue('N'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // VAT rate [%]
    // $template_active_worksheet->setCellValue('O'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // <predkontacia>
}

if ($debug_only) print("</pre>\n");


if (! $debug_only) {
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="processed_orders.xls"');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

    $writer = IOFactory::createWriter($template_spreadsheet, 'Xls');
    $writer->save('php://output');
}

?>
