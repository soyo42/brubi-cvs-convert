<?php
$debug_only = boolval(isset($_POST['debug-only']));
if (! $debug_only) {
    // fallback to GET request
    $debug_only = boolval(isset($_REQUEST['debug']));
}

require 'util.php';

// if ($debug_only) {
//     print('debug yes!');
// } else {
//     print('debug no!');
// }       
// exit;


if(isset($_POST['submit'])) {
    if (isset($_FILES['userfile'])) {
        $countfiles = count($_FILES['userfile']['name']);
        if ($countfiles != 3) {
            exit("potrebujeme presne 3 subory, ziskali sme $countfiles}");
        }

        check_uploaded_file_names($_FILES['userfile']['name']);

        $idx = 0;
        foreach ($_FILES['userfile']['name'] as $fileName) {
            if ($debug_only) printf("received file name: %s --> %s<br/>\n", $fileName, $_FILES['userfile']['tmp_name'][$idx]);
            $idx += 1;
        }
        

        $orders_file = $_FILES['userfile']['tmp_name'][0];
        $items_file = $_FILES['userfile']['tmp_name'][1];
        $refund_statements_file = $_FILES['userfile']['tmp_name'][2];
    }
} else {
    if ($debug_only) print("<b>DEMO - with static files</b><br>\n");
    $orders_file = './EtsySoldOrders2023-3.csv';
    $items_file = './EtsySoldOrderItems2023-3.csv';
    $refund_statements_file = './etsy_statement_2023_3.csv';
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
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;

$sadzby_csv = './sadzby_DPH_s_predkontaciou.csv';
$template_xls = './result_template.xls';

// load sadzby via csv file
$sadzby_reader = new CsvReader();
$sadzby_reader->setDelimiter(';')
    ->setEnclosure('"')
    ->setSheetIndex(0);
$sadzby_spreadsheet = $sadzby_reader->load($sadzby_csv);

// load template xls
$template_spreadsheet = IOFactory::load($template_xls);
$template_active_worksheet = $template_spreadsheet->getActiveSheet();



if ($debug_only) print("<pre>\n");


// read sadzby + staty
$sadzby_active_worksheet = $sadzby_spreadsheet->getActiveSheet();
$SADZBY = sadzby_to_map($sadzby_active_worksheet);
if ($debug_only) {
    print("--sadzby--\n");
    foreach ($SADZBY as $key => $value) {
        printf("@@ %-15s [%2s | %4s] -- %s\n", $key, $value['C'], $value['D'], $value['country_isocode']);
    }
}


if ($debug_only) print("----\n");
// load csv files (delimiter = ',')
$reader = new CsvReader();
$reader->setDelimiter(',')
    ->setEnclosure('"')
    ->setSheetIndex(0);

// process orders, items, refunds
$orders_active_worksheet = $reader->load($orders_file)->getActiveSheet();
$order_items_active_worksheet = $reader->load($items_file)->getActiveSheet();
$refund_statements_worksheet = $reader->load($refund_statements_file)->getActiveSheet();

$items_map = order_items_to_map($order_items_active_worksheet);
$refund_statements_info = analyze_refund_statements($refund_statements_worksheet);
if ($debug_only) printf("--- refund stats: %s ---\n\n", $refund_statements_info->statistics);


$row = 0;
$partial_refund_rows = [];
$orders_coliterator = new ColumnCellIterator($orders_active_worksheet, 'A', 2);
foreach ($orders_coliterator as $cell) {
    $row = $cell->getRow();
    $country_name = $orders_active_worksheet->getCell('O'.$row)->getValue();
    $order_id = $orders_active_worksheet->getCell('B'.$row)->getValue();
    
    if ($debug_only) printf("--- source: %03d %s orderId[%s] --- ", $row, $cell->getValue(), $order_id);

    $target_row = $row + 2;
    prepare_template_row($template_active_worksheet, $target_row);

    $cell_date = DateTime::createFromFormat('m/d/y', $cell->getValue());

    if ($debug_only) {
        printf("target: type[%s] -> %s\n",
               $template_active_worksheet->getCell('A'.($target_row-1))->getDataType(),
               $cell_date->format('d.m.Y')
        );
    }

    // resolve sadzba
    if (array_key_exists($country_name, $SADZBY)) {
        $sadzba = $SADZBY[$country_name];
    } else {
        $sadzba = $SADZBY['DEFAULT'];
    }

    // resolve item name
    if (array_key_exists($order_id, $items_map)) {
        $item_name = $items_map[$order_id];
    } else {
        $item_name = '?';
    }

    // process partial refunds in now
    $comment = null;
    $price = $orders_active_worksheet->getCell('X'.$row)->getValue();
    if (array_key_exists($order_id, $refund_statements_info->partial_refunds_now)) {
        $comment = 'ciastocny refund';
        $refund_price = $refund_statements_info->partial_refunds_now[$order_id][0];
        if ($debug_only) printf("partial refund @ $order_id: $price $refund_price\n");
        $price += $refund_price;
        $partial_refund_rows []= $target_row;
    }
    
    $template_active_worksheet->setCellValue('A'.$target_row, XlsDate::PHPToExcel($cell_date));  // date
    $template_active_worksheet->setCellValue('B'.$target_row, $order_id);  // order no.
    $template_active_worksheet->setCellValue('C'.$target_row, $orders_active_worksheet->getCell('D'.$row)->getValue());  // full name
    $template_active_worksheet->setCellValue('D'.$target_row, $item_name);  // full name
    $template_active_worksheet->setCellValue('E'.$target_row, $orders_active_worksheet->getCell('G'.$row)->getValue());  // number of items
    $template_active_worksheet->setCellValue('F'.$target_row, $orders_active_worksheet->getCell('J'.$row)->getValue());  // street name
    $template_active_worksheet->setCellValue('G'.$target_row, $orders_active_worksheet->getCell('L'.$row)->getValue());  // city
    $template_active_worksheet->setCellValue('H'.$target_row, $orders_active_worksheet->getCell('N'.$row)->getValue());  // zip code
    $template_active_worksheet->setCellValue('I'.$target_row, $orders_active_worksheet->getCell('O'.$row)->getValue());  // country
    $template_active_worksheet->setCellValue('J'.$target_row, $orders_active_worksheet->getCell('P'.$row)->getValue());  // currency
    $template_active_worksheet->setCellValue('K'.$target_row, $sadzba['E']);  // OSS - <druh dodani hlavicka>
    $template_active_worksheet->setCellValue('L'.$target_row, $sadzba['country_isocode']);  // OSS - country
    $template_active_worksheet->setCellValue('M'.$target_row, $price);  // OSS - base rate
    $template_active_worksheet->setCellValue('N'.$target_row, $sadzba['C']);  // VAT rate [%]
    $template_active_worksheet->setCellValue('O'.$target_row, $sadzba['D']);  // <predkontacia>
    if ($comment != null) {
        $template_active_worksheet->setCellValue('P'.$target_row, $comment);  // comment
    }
}

// color modified rows (changed through partial refund)
foreach($partial_refund_rows as $key => $partial_refund_row) {
    change_row_background_color($template_active_worksheet, $partial_refund_row, 'fff46d');
}


// insert empty row before past refunds
$target_row += 1;
$template_active_worksheet->insertNewRowBefore($target_row);


// process full refunds in past
foreach ($refund_statements_info->full_refunds_in_past as $order_id => $record) {
    $target_row += 1;
    prepare_template_row($template_active_worksheet, $target_row);
    if ($debug_only) printf("refund [$order_id] date: %s -> $record[0]\n", $record[1]->format('d.m.Y'));
    fill_refund_row($template_active_worksheet, $target_row, $record, $order_id, 'plny refund');
    change_row_background_color($template_active_worksheet, $target_row, 'ff0000');
    
}

// process partial refunds in past
foreach ($refund_statements_info->partial_refunds_in_past as $order_id => $record) {
    $target_row += 1;
    prepare_template_row($template_active_worksheet, $target_row);
    if ($debug_only) printf("partial refund [$order_id] date: %s -> $record[0]\n", $record[1]->format('d.m.Y'));
    fill_refund_row($template_active_worksheet, $target_row, $record, $order_id, 'ciastocny refund');
    change_row_background_color($template_active_worksheet, $target_row, 'ff6d6d');
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
