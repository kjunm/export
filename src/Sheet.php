<?php
/**
 * Created by PhpStorm.
 * User: wim
 * Date: 2018/10/24
 * Time: 17:37
 */
namespace wim\export;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Sheet
{
    private $spreadSheet;

    public $templatePath;

    /**
     * Sheet constructor.
     * @param string $path
     */
    public function __construct($path='')
    {
        if(file_exists($path)){
            $this->templatePath = $path;
        }
    }

    /**
     * set spreadSheet
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function setSpreadSheet()
    {
        if(!empty($this->templatePath)){
            $this->spreadSheet = IOFactory::load($this->templatePath);
        }else{
            $this->spreadSheet = new Spreadsheet();
        }
    }

    /**
     * get spreadSheet
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function getSpreadSheet()
    {
        if($this->spreadSheet == null){
            $this->spreadSheet = $this->setSpreadSheet();
        }
        return $this->spreadSheet;
    }

    /**
     * get active sheet
     */
    public function getActiveSheet()
    {
        if($this->spreadSheet == null){
            $this->spreadSheet = $this->setSpreadSheet();
        }
        return $this->spreadSheet->getActiveSheet();
    }


    /**
     * delete spreadSheet
     */
    public function deleteSpreadSheet()
    {
        $this->getSpreadSheet()->disconnectWorksheets();
    }

    /**
     * set title
     * @param $title 工作表名称
     * @throws Exception
     */
    public function setTitle($title)
    {
        $this->getActiveSheet()->setTitle($title);
    }

    /**
     * set header
     * @param array $headers 表头
     */
    public function setHeader(array $headers)
    {
        $this->getActiveSheet()->fromArray($headers);
    }

    /**
     * set content
     */
    public function setContent()
    {

    }










}