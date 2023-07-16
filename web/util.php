<?php

use PhpOffice\PhpSpreadsheet\Worksheet\ColumnCellIterator;
use PhpOffice\PhpSpreadsheet\Worksheet\RowCellIterator;
use PhpOffice\PhpSpreadsheet\Shared\Date as XlsDate;
use PhpOffice\PhpSpreadsheet\Style\Fill;
require 'RefundStatementInfo.class.php';


define('FILE_PATTERNS',
       [
           0 => '/^EtsySoldOrders([0-9]{4})-([0-9]{1,2}).csv$/',
           1 => '/^EtsySoldOrderItems([0-9]{4})-([0-9]{1,2}).csv$/',
           2 => '/^etsy_statement_([0-9]{4})_([0-9]{1,2}).csv$/'
       ]
);


function prepare_template_row($worksheet, $target_row) {
    $worksheet->insertNewRowBefore($target_row);
    $worksheet->getRowDimension($target_row)->setRowHeight(15);
    $worksheet->getStyle('A'.$target_row)
                              ->getNumberFormat()
                              ->setFormatCode('dd.mm.YYYY');

    $worksheet->getStyle('M'.$target_row)
              ->getNumberFormat()
              ->setFormatCode('0.00');
}


function change_row_background_color($worksheet, $target_row, $colorRgb) {
    $rowIterator = new RowCellIterator($worksheet, $target_row, 'A', 'P');
    foreach ($rowIterator as $cell) {
        $column = $cell->getColumn();
        $worksheet->getStyle($column.$target_row)
                  ->getFill()
                  ->setFillType(Fill::FILL_SOLID)
                  ->getStartColor()
                  ->setARGB($colorRgb);
    }
}


function check_uploaded_file_names($fileNames) {
    $error_messages = '';
    $used_years = [];
    $used_months = [];
    
    $idx = 0;
    foreach ($fileNames as $fileName) {
        if (array_key_exists($idx, FILE_PATTERNS)) {
            $pattern = FILE_PATTERNS[$idx];
            $found = preg_match($pattern, $fileName, $matches);
            if (! $found) {
                $error_messages .= sprintf('nespravny subor [%s]\n&nbsp;expected pattern: %s\n', $fileName, $pattern);
            } else {
                $used_years []= $matches[1];
                $used_months []= $matches[2];
            }
        } else {
            $error_messages .= 'unexpected index -> {$idx}\n';
        }
        $idx += 1;
    }

    foreach ($used_years as $key => $year) {
        $month = $used_months[$key];
        if ($year != $used_years[0]) {
            $error_messages .= sprintf('nespravny rok [%s]: %s\n &nbsp;ocakavany: %s\n', $fileNames[$key], $year, $used_years[0]);
        }
        if ($month != $used_months[0]) {
            $error_messages .= sprintf('nespravny mesiac [%s]: %s\n &nbsp;ocakavany: %s\n', $fileNames[$key], $month, $used_months[0]);
        }
    }

    if ($error_messages != '') {
        exit(str_replace('\n', "<br/>\n",  $error_messages));
    }
}


function fill_refund_row($worksheet, $target_row, $record, $order_id, $comment) {
    $worksheet->setCellValue('A'.$target_row, XlsDate::PHPToExcel($record[1]));  // date
    $worksheet->setCellValue('B'.$target_row, $order_id);  // order no.
    $worksheet->setCellValue('M'.$target_row, $record[0]);  // OSS - base rate
    $worksheet->setCellValue('P'.$target_row, $comment);  // comment

}


function sadzby_to_map($sadzby_active_worksheet) {
    $SADZBY = array();
    for ($i = 1; true; $i++) {
        $country = $sadzby_active_worksheet->getCell("A$i")->getValue();
        $country_isocode = $sadzby_active_worksheet->getCell("B$i")->getValue();
        if ($country == null) {
            break;
        }
        $c = $sadzby_active_worksheet->getCell("C$i")->getValue();
        $d = $sadzby_active_worksheet->getCell("D$i")->getValue();
        $e = $sadzby_active_worksheet->getCell("E$i")->getValue();
        // print("country: $country [$country_isocode] \n");
        $SADZBY[$country] = array('C' => $c, 'D' => $d, 'E' => $e, 'country_isocode' => $country_isocode);
        // $width = 15 - mb_strlen($country) + strlen($country);
        // printf("%-{$width}s :: $c -- $d\n", $country);
    }
    return $SADZBY;
}


function order_items_to_map($items_active_worksheet) {
    $ITEMS = array();
    $colIterator = new ColumnCellIterator($items_active_worksheet, 'B', 2);
    foreach ($colIterator as $cell) {
        $row = $cell->getRow();
        $order_id = $items_active_worksheet->getCell('Y'.($row))->getValue();
        $ITEMS[$order_id] = $cell->getValue();
    }
    return $ITEMS;
}


function analyze_refund_statements($refund_statements_worksheet) {
    $info = new RefundStatementInfo();
    $REFUND_PATTERN = '/^Refund to buyer for Order #([0-9]+)$/';
    $PARTIAL_REFUND_PATTERN = '/^Partial refund to buyer for Order #([0-9]+)$/';
    $PAYMENT_PATTERN = '/^Payment for Order #([0-9]+)$/';

    $refunds = [];
    $partial_refunds = [];
    $payments = [];

    $colIterator = new ColumnCellIterator($refund_statements_worksheet, 'C', 2);
    foreach ($colIterator as $cell) {
        $row = $cell->getRow();
        $title = $cell->getValue();
        $raw_price = $refund_statements_worksheet->getCell('F'.($row))->getValue();

        $date = $refund_statements_worksheet->getCell('A'.($row))->getValue();
        if (preg_match($REFUND_PATTERN, $title, $matches)) {
            $refunds[$matches[1]] = _build_record($title, $raw_price, $date);
        } else if (preg_match($PARTIAL_REFUND_PATTERN, $title, $matches)) {
            $partial_refunds[$matches[1]] = _build_record($title, $raw_price, $date);
        } else if (preg_match($PAYMENT_PATTERN, $title, $matches)) {
            $payments[$matches[1]] = _build_record($title, $raw_price, $date);
        }
    }

    $info->statistics = sprintf("ref: %s, paref: %s, pay: %s", sizeof($refunds), sizeof($partial_refunds), sizeof($payments));

    // process full refunds
    foreach ($refunds as $order_id => $record) {
        if (array_key_exists($order_id, $payments)) {
            if ($payments[$order_id][0] != -$record[0]) {
                // this went probably wrong!
                print_r($payments[$order_id]);
                print_r($record);
                exit("full refund price mismatch: orderId($order_id), (-){$payments[$order_id][0]} <> $record[0]");
            }
        } else {
            // full refund in past
            $info->full_refunds_in_past[$order_id] = $record;
        }
    }
    // process partial refunds
    foreach ($partial_refunds as $order_id => $record) {
        if (array_key_exists($order_id, $payments)) {
            // partial refund now
            $info->partial_refunds_now[$order_id] = $record;
        } else {
            // partial refund in past
            $info->partial_refunds_in_past[$order_id] = $record;
        }
    }

    return $info;
}


function _build_record($title, $raw_price, $date) {
    $PRICE_PATTERN = '/^-?[0-9.]+$/';
    $price_candidate = filter_var($raw_price, FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);
    // printf("price candidate: $price_candidate [$title] \n");
    if (preg_match($PRICE_PATTERN, $price_candidate)) {
        $price = floatval($price_candidate);
    } else {
        exit("failed to parse price in refund statement [$title]: '$raw_price'");
    }

    $nice_date = DateTime::createFromFormat('d M, Y', $date);
    #printf("raw date: $date ---> %s\n", $nice_date);
    return [$price, $nice_date];
}

?>
