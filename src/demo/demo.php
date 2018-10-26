<?php
/**
 * Created by PhpStorm.
 * User: wim
 * Date: 2018/10/23
 * Time: 14:31
 */
namespace wim\export\demo;

use PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_once __DIR__."/../../vendor/autoload.php";

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','客流');
$sheet->setCellValue('A2','客流');
$sheet->setCellValue('A3','客流');
$sheet->setCellValue('A4','客流');
$sheet->mergeCells('A1:A5');
//$dateTimeNow = time();
//$excelDateValue = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $dateTimeNow );
//// Set cell A6 with the Excel date/time value
//$spreadsheet->getActiveSheet()->setCellValue(
//    'A6',
//    $excelDateValue
//);
//// Set the number format mask so that the excel timestamp will be displayed as a human-readable date/time
//$spreadsheet->getActiveSheet()->getStyle('A6')
//    ->getNumberFormat()
//    ->setFormatCode(
//        'yyyy-mm-dd hh:mm:ss'
//    );
//$spreadsheet->getActiveSheet()->setCellValueByColumnAndRow(1, 5, 'PhpSpreadsheet');
//
//
//// Set value binder
//\PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );
//
//// Create new Spreadsheet object
//$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
//
//// ...
//// Add some data, resembling some different data types
//$spreadsheet->getActiveSheet()->setCellValue('A4', 'Percentage value:');
//// Converts the string value to 0.1 and sets percentage cell style
//$spreadsheet->getActiveSheet()->setCellValue('B4', '10%');
//
//$spreadsheet->getActiveSheet()->setCellValue('A5', 'Date/time value:');
//// Converts the string value to an Excel datestamp and sets the date format cell style
//$spreadsheet->getActiveSheet()->setCellValue('B5', '21 December 1983');


$writer = new Xlsx($spreadsheet);

$filename = 'helloworld.xlsx';
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$filename.'"');
header('Cache-Control: max-age=0');
$writer->save('php://output');