<?php
/**
 * Created by PhpStorm.
 * User: wim
 * Date: 2018/11/15
 * Time: 9:40
 */

namespace wim\export\demo;

use wim\export\Sheet;
use wim\export\SheetBuilder;

include '../../vendor/autoload.php';

class demo extends SheetBuilder
{
    public $sheet;

    public function __construct()
    {
        $this->sheet = new Sheet();

    }

    public function buildTitle()
    {
        $this->sheet->setTitle('12123123131');
    }

    public function buildHeaders()
    {
        $this->sheet->setHeader(['a','b','c']);
    }

    public function buildContent()
    {
        $this->sheet->getActiveSheet()->fromArray([1,2,3],null,'A3');
    }

    public function getResult()
    {
        $this->buildTitle();
        $this->buildHeaders();
        $this->buildContent();

        $filename = '123123.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->sheet->getSpreadSheet(), 'Xlsx');
        $writer->save('php://output');
    }
}

$builder = new demo();
$builder->getResult();
