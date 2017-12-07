<?php
namespace Ellumilel;

use Ellumilel\DocProps\App;
use Ellumilel\DocProps\Core;
use Ellumilel\Helpers\ExcelHelper;
use Ellumilel\Rels\Relationships;
use Ellumilel\Xl\SharedStrings;
use Ellumilel\Xl\Styles;
use Ellumilel\Xl\Workbook;
use Ellumilel\Xl\Worksheets\SheetXml;

/**
 * Class ExcelWriter
 * @package Ellumilel
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class ExcelWriter
{
    /** @var array */
    protected $sheets = [];
    /** @var array */
    protected $sharedStrings = [];
    /** @var int */
    protected $sharedStringCount = 0;
    /** @var array */
    protected $tempFiles = [];
    /** @var array */
    protected $cellFormats = [];
    /** @var array */
    protected $cellTypes = [];
    /** @var string  */
    protected $currentSheet = '';
    /** @var null */
    protected $tmpDir = null;
    /** @var null */
    protected $fileName = null;
    /** @var Core */
    protected $core;
    /** @var App */
    protected $app;
    /** @var Workbook */
    protected $workbook;
    /** @var SheetXml */
    protected $sheetXml;

    /**
     * ExcelWriter constructor.
     * @throws \Exception
     */
    public function __construct()
    {
        if (!class_exists('ZipArchive')) {
            throw new \Exception('ZipArchive not found');
        }
        if (!ini_get('date.timezone')) {
            date_default_timezone_set('UTC');
        }
        $this->addCellFormat($cell_format = 'GENERAL');
        $this->core = new Core();
        $this->app = new App();
        $this->workbook = new Workbook();
        $this->sheetXml = new SheetXml();
    }

    /**
     * @param string $author
     */
    public function setAuthor($author)
    {
        $this->core->setAuthor($author);
    }

    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                if (file_exists($tempFile)) {
                    unlink($tempFile);
                }
            }
        }
    }

    /**
     * @param $dir
     */
    public function setTmpDir($dir)
    {
        $this->tmpDir = $dir;
    }

    /**
     * Set output filename: yourFileName.xlsx
     *
     * @param string $fileName
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
    }

    /**
     * Return tmpFileName
     * @return string
     */
    protected function tempFilename()
    {
        $tmpDir = is_null($this->tmpDir) ? sys_get_temp_dir() : $this->tmpDir;
        $filename = @tempnam($tmpDir, "excelWriter_");
        $this->tempFiles[] = $filename;

        return $filename;
    }

    /**
     * @param bool $headers
     */
    public function writeToStdOut($headers = true)
    {
        if (empty($this->tmpDir)) {
            $tempFile = $this->tempFilename().'.xlsx';
        } else {
            $tempFile = $this->fileName;
        }

        $this->writeToFile($tempFile);
        if (file_exists($tempFile)) {
            if ($headers) {
                header('Content-Description: File Transfer');
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename="'.basename($tempFile).'"');
                header('Expires: 0');
                header('Cache-Control: must-revalidate');
                header('Pragma: public');
                header('Content-Length: '.filesize($tempFile));
            }
            readfile($tempFile);
        }
    }

    /**
     * @return string
     */
    public function writeToString()
    {
        $tempFile = $this->tempFilename();
        $this->writeToFile($tempFile);
        $string = file_get_contents($tempFile);

        return $string;
    }

    /**
     * @param string $filename
     */
    public function writeToFile($filename)
    {
        $zip = new \ZipArchive();
        foreach ($this->sheets as $sheetName => $sheet) {
            $this->finalizeSheet($sheetName);
        }
        $this->checkAndUnlink($zip, $filename);
        $this->workbook->setSheet($this->sheets);

        $contentTypes = new ContentTypes(!empty($this->sharedStrings));
        $contentTypes->setSheet($this->sheets);

        $rel = new Relationships(!empty($this->sharedStrings));
        $rel->setSheet($this->sheets);
        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", $this->app->buildAppXML());
        $zip->addFromString("docProps/core.xml", $this->core->buildCoreXML());
        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", $rel->buildRelationshipsXML());
        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            /** @var Sheet $sheet */
            $zip->addFile($sheet->getFilename(), "xl/worksheets/".$sheet->getXmlName());
        }
        if (!empty($this->sharedStrings)) {
            $zip->addFile($this->writeSharedStringsXML(), "xl/sharedStrings.xml");
        }
        $zip->addFromString("xl/workbook.xml", $this->workbook->buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");
        $zip->addFromString("[Content_Types].xml", $contentTypes->buildContentTypesXML());
        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", $rel->buildWorkbookRelationshipsXML());
        $zip->close();
    }

    /**
     * @param string $sheetName
     */
    protected function initializeSheet($sheetName)
    {
        if ($this->currentSheet == $sheetName || isset($this->sheets[$sheetName])) {
            return;
        }
        $sheetFilename = $this->tempFilename();
        $sheetXmlName = 'sheet' . (count($this->sheets) + 1).".xml";
        $sheetObj = new Sheet();
        $sheetObj
            ->setFilename($sheetFilename)
            ->setSheetName($sheetName)
            ->setXmlName($sheetXmlName)
            ->setWriter(new Writer($sheetFilename))
        ;
        $this->sheets[$sheetName] = $sheetObj;
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $selectedTab = count($this->sheets) == 1 ? 'true' : 'false';
        $maxCell = ExcelHelper::xlsCell(ExcelHelper::EXCEL_MAX_ROW, ExcelHelper::EXCEL_MAX_COL);
        $sheet->getWriter()->write($this->sheetXml->getXml());
        $sheet->getWriter()->write($this->sheetXml->getWorksheet());
        $sheet->getWriter()->write($this->sheetXml->getSheetPr());
        $sheet->setMaxCellTagStart($sheet->getWriter()->fTell());
        $sheet->getWriter()->write($this->sheetXml->getDimension($maxCell));
        $sheet->setMaxCellTagEnd($sheet->getWriter()->fTell());
        $sheet->getWriter()->write($this->sheetXml->getSheetViews($selectedTab));
        $sheet->getWriter()->write($this->sheetXml->getCools());
        $sheet->getWriter()->write('<sheetData>');
    }

    /**
     * @param $cellFormat
     *
     * @return string
     */
    private function determineCellType($cellFormat)
    {
        $cellFormat = str_replace("[RED]", "", $cellFormat);
        if ($cellFormat == 'GENERAL') {
            return 'string';
        }
        if ($cellFormat == '0') {
            return 'numeric';
        }
        if ($cellFormat == '@') {
            return '@';
        }
        $checkArray = [
            'datetime' => [
                "/[H]{1,2}:[M]{1,2}/",
                "/[M]{1,2}:[S]{1,2}/",
            ],
            'numeric' => [
                "/0/",
            ],
            'date' => [
                "/[YY]{2,4}/",
                "/[D]{1,2}/",
                "/[M]{1,2}/",
            ],
            'currency' => [
                "/$/",
            ],
            'percent' => [
                "/%/",
            ],
        ];
        foreach ($checkArray as $type => $item) {
            foreach ($item as $prMatch) {
                if (preg_match($prMatch, $cellFormat)) {
                    return $type;
                }
            }
        }

        return 'string';
    }

    /**
     * backwards compatibility
     *
     * @param $cellFormat
     *
     * @return int|mixed
     */
    private function addCellFormat($cellFormat)
    {
        $cellFormat = strtoupper($this->getCellFormat($cellFormat));
        $position = array_search($cellFormat, $this->cellFormats, $strict = true);
        if ($position === false) {
            $position = count($this->cellFormats);
            $this->cellFormats[] = ExcelHelper::escapeCellFormat($cellFormat);
            $this->cellTypes[] = $this->determineCellType($cellFormat);
        }

        return $position;
    }

    /**
     * @link https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats(v=office.15).aspx
     *
     * @param string $cellFormat
     *
     * @return string
     */
    private function getCellFormat($cellFormat)
    {
        $formatArray = [
            'string' => 'GENERAL',
            'text' => '@',
            'integer' => '0',
            'float_with_sep' => '#,##0.00',
            'float' => '0.00',
            'date' => 'YYYY-MM-DD',
            'datetime' => 'YYYY-MM-DD HH:MM:SS',
            'dollar' => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
            'money' => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
            'euro' => '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]',
            'rub' => '#,##0.00 [$₽-419];[Red]-#,##0.00 [$₽-419]',
            'NN' => 'DDD',
            'NNN' => 'DDDD',
            'NNNN' => 'DDDD", "',
        ];

        if (array_key_exists($cellFormat, $formatArray)) {
            return $formatArray[$cellFormat];
        }
        return $cellFormat;
    }

    /**
     * @param string $sheetName
     * @param array $headerTypes
     * @param bool $suppressRow
     */
    public function writeSheetHeader($sheetName, array $headerTypes, $suppressRow = false)
    {
        if (empty($sheetName) || empty($headerTypes) || !empty($this->sheets[$sheetName])) {
            return;
        }
        $this->initializeSheet($sheetName);
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $sheet->setColumns([]);
        foreach ($headerTypes as $val) {
            $sheet->setColumn($this->addCellFormat($val));
        }
        if (!$suppressRow) {
            $this->writeRowHeader($sheet, array_keys($headerTypes));
            $sheet->increaseRowCount();
        }
        $this->currentSheet = $sheetName;
    }

    /**
     * @param Sheet $sheet
     * @param array $headerRow
     */
    private function writeRowHeader(Sheet $sheet, $headerRow)
    {
        $sheet->getWriter()->write(
            '<row collapsed="false" customFormat="false" 
                customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.(1).'">'
        );
        foreach ($headerRow as $k => $v) {
            $this->writeCell($sheet->getWriter(), 0, $k, $v, $cell_format_index = '0');
        }
        $sheet->getWriter()->write('</row>');
    }

    /**
     * @param string $sheetName
     * @param array $row
     */
    public function writeSheetRow($sheetName, array $row)
    {
        if (empty($sheetName) || empty($row)) {
            return;
        }
        $this->initializeSheet($sheetName);
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $columns = $sheet->getColumns();
        if (empty($columns)) {
            $sheet->setColumns(array_fill(0, count($row), '0'));
        }
        $sheet->getWriter()->write(
            '<row collapsed="false" customFormat="false" customHeight="false" 
            hidden="false" ht="12.1" outlineLevel="0" r="'.($sheet->getRowCount() + 1).'">'
        );
        $column_count = 0;
        $sheetColumns = $sheet->getColumns();
        foreach ($row as $k => $v) {
            $this->writeCell(
                $sheet->getWriter(),
                $sheet->getRowCount(),
                $column_count,
                $v,
                $sheetColumns[$column_count]
            );
            $column_count++;
        }
        $sheet->getWriter()->write('</row>');
        $sheet->increaseRowCount();
        $this->currentSheet = $sheetName;
    }

    /**
     * @param string $sheetName
     */
    protected function finalizeSheet($sheetName)
    {
        if (empty($sheetName) || ($this->sheets[$sheetName] instanceof Sheet &&
                $this->sheets[$sheetName]->getFinalized()
            )
        ) {
            return;
        }

        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $sheet->getWriter()->write('</sheetData>');
        $mergeCells = $sheet->getMergeCells();
        if (!empty($mergeCells)) {
            $sheet->getWriter()->write($this->sheetXml->getMergeCells($mergeCells));
        }
        $sheet->getWriter()->write($this->sheetXml->getPrintOptions());
        $sheet->getWriter()->write($this->sheetXml->getPageMargins());
        $sheet->getWriter()->write($this->sheetXml->getPageSetup());
        $sheet->getWriter()->write($this->sheetXml->getHeaderFooter());
        $sheet->getWriter()->write('</worksheet>');
        $maxCell = ExcelHelper::xlsCell($sheet->getRowCount() - 1, count($sheet->getColumns()) - 1);
        $maxCellTag = $this->sheetXml->getDimension($maxCell);
        $paddingLength = $sheet->getMaxCellTagEnd() - $sheet->getMaxCellTagStart() - strlen($maxCellTag);
        $sheet->getWriter()->fSeek($sheet->getMaxCellTagStart());
        $sheet->getWriter()->write($maxCellTag.str_repeat(" ", $paddingLength));
        $sheet->getWriter()->close();
        $sheet->setFinalized(true);
    }

    /**
     * @param string $sheetName
     * @param int $startCellRow
     * @param int $startCellColumn
     * @param int $endCellRow
     * @param int $endCellColumn
     */
    public function markMergedCell($sheetName, $startCellRow, $startCellColumn, $endCellRow, $endCellColumn)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->getFinalized()) {
            return;
        }
        $this->initializeSheet($sheetName);
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $startCell = ExcelHelper::xlsCell($startCellRow, $startCellColumn);
        $endCell = ExcelHelper::xlsCell($endCellRow, $endCellColumn);
        $sheet->setMergeCells($startCell.":".$endCell);
    }

    /**
     * @param array $data
     * @param string $sheetName
     * @param array $headerTypes
     */
    public function writeSheet(array $data, $sheetName = '', array $headerTypes = [])
    {
        $sheetName = empty($sheetName) ? 'Sheet1' : $sheetName;
        $data = empty($data) ? [['']] : $data;
        if (!empty($headerTypes)) {
            $this->writeSheetHeader($sheetName, $headerTypes);
        }
        foreach ($data as $i => $row) {
            $this->writeSheetRow($sheetName, $row);
        }
        $this->finalizeSheet($sheetName);
    }

    /**
     * @param Writer $file
     * @param int $rowNumber
     * @param int $columnNumber
     * @param mixed $value
     * @param $cellIndex
     */
    protected function writeCell(Writer $file, $rowNumber, $columnNumber, $value, $cellIndex)
    {
        $cellType = $this->cellTypes[$cellIndex];
        $cellName = ExcelHelper::xlsCell($rowNumber, $columnNumber);
        $cell = $this->sheetXml->getCell($cellName, $cellIndex, $cellType, $value);
        if ($cell === false) {
            $file->write(
                '<c r="'.$cellName.'" s="'.$cellIndex.'" t="s"><v>'.ExcelHelper::xmlspecialchars(
                    $this->setSharedString($value)
                ).'</v></c>'
            );
        } else {
            $file->write($cell);
        }
    }

    /**
     * @return string
     */
    protected function writeStylesXML()
    {
        $temporaryFilename = $this->tempFilename();
        $file = new Writer($temporaryFilename);
        $styles = new Styles();
        $styles->setCellFormats($this->cellFormats);
        $file->write($styles->buildStylesXML());

        return $temporaryFilename;
    }

    /**
     * @param $v
     *
     * @return int|mixed
     */
    protected function setSharedString($v)
    {
        if (isset($this->sharedStrings[$v])) {
            $stringValue = $this->sharedStrings[$v];
        } else {
            $stringValue = count($this->sharedStrings);
            $this->sharedStrings[$v] = $stringValue;
        }
        $this->sharedStringCount++;

        return $stringValue;
    }

    /**
     * @return string
     */
    protected function writeSharedStringsXML()
    {
        $tempFilename = $this->tempFilename();
        $file = new Writer($tempFilename);
        $sharedStrings = new SharedStrings($this->sharedStringCount, $this->sharedStrings);
        $file->write($sharedStrings->buildSharedStringsXML());
        $file->close();

        return $tempFilename;
    }

    /**
     * @param \ZipArchive $zip
     * @param string $filename
     */
    private function checkAndUnlink(\ZipArchive $zip, $filename)
    {
        if (file_exists($filename) && is_writable($filename)) {
            unlink($filename);
        }
        if (empty($this->sheets) || !$zip->open($filename, \ZipArchive::CREATE)) {
            throw new \RuntimeException(
                "Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined or unable to create zip."
            );
        }
    }
}
