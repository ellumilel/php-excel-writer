<?php
namespace Ellumilel;

use Ellumilel\DocProps\App;
use Ellumilel\DocProps\Core;
use Ellumilel\Rels\Relationships;
use Ellumilel\Xl\SharedStrings;
use Ellumilel\Xl\Styles;
use Ellumilel\Xl\Workbook;

/**
 * Class ExcelWriter
 * @package Ellumilel
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class ExcelWriter
{
    /**
     * @link http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    const EXCEL_MAX_ROW = 1048576;
    const EXCEL_MAX_RANGE = 2147483647;
    const EXCEL_MAX_COL = 16384;

    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/officeDocument/2006';

    /** @var string */
    protected $author ='Unknown Author';
    /** @var array */
    protected $sheets = [];
    /** @var array */
    protected $sharedStrings = [];//unique set
    /** @var int */
    protected $sharedStringCount = 0;//count of non-unique references to the unique set
    /** @var array */
    protected $tempFiles = [];
    /** @var array */
    protected $cellFormats = [];//contains excel format like YYYY-MM-DD HH:MM:SS
    /** @var array */
    protected $cellTypes = [];//contains friendly format like datetime
    /** @var string  */
    protected $currentSheet = '';
    /** @var null */
    protected $tmpDir = null;
    /** @var Core */
    protected $core;
    /** @var App */
    protected $app;
    /** @var Workbook */
    protected $workbook;

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
            //using date functions can kick out warning if this isn't set
            date_default_timezone_set('UTC');
        }
        $this->addCellFormat($cell_format = 'GENERAL');
        $this->core = new Core();
        $this->app = new App();
        $this->workbook = new Workbook();
    }

    /**
     * @param string $author
     */
    public function setAuthor($author = '')
    {
        $this->author = $author;
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
     * Return tmpFileName
     * @return string
     */
    protected function tempFilename()
    {
        $tmpDir = is_null($this->tmpDir) ? sys_get_temp_dir() : $this->tmpDir;
        $filename = tempnam($tmpDir, "exlsWriter_");
        $this->tempFiles[] = $filename;

        return $filename;
    }

    public function writeToStdOut()
    {
        $tempFile = $this->tempFilename();
        $this->writeToFile($tempFile);
        readfile($tempFile);
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
        foreach ($this->sheets as $sheetName => $sheet) {
            $this->finalizeSheet($sheetName);
        }
        if (file_exists($filename) && is_writable($filename)) {
            unlink($filename);
        }

        $zip = new \ZipArchive();
        if (empty($this->sheets) || !$zip->open($filename, \ZipArchive::CREATE)) {
            self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined or unable to create zip.");
            return;
        }

        $this->workbook->setSheet($this->sheets);

        $contentTypes = new ContentTypes(!empty($this->sharedStrings));
        $contentTypes->setSheet($this->sheets);

        $rels = new Relationships(!empty($this->sharedStrings));
        $rels->setSheet($this->sheets);

        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", $this->app->buildAppXML());
        $zip->addFromString("docProps/core.xml", $this->core->buildCoreXML());
        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", $rels->buildRelationshipsXML());
        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            /** @var Sheet $sheet */
            $zip->addFile($sheet->getFilename(), "xl/worksheets/".$sheet->getXmlName());
        }
        if (!empty($this->sharedStrings)) {
            $zip->addFile(
                $this->writeSharedStringsXML(),
                "xl/sharedStrings.xml"
            );
        }
        $zip->addFromString("xl/workbook.xml", $this->workbook->buildWorkbookXML());
        $zip->addFile(
            $this->writeStylesXML(),
            "xl/styles.xml"
        );
        $zip->addFromString("[Content_Types].xml", $contentTypes->buildContentTypesXML());
        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", $rels->buildWorkbookRelationshipsXML());
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
        $selectedTab = count($this->sheets) == 1 ? 'true' : 'false';//only first sheet is selected
        $maxCell = ExcelWriter::xlsCell(self::EXCEL_MAX_ROW, self::EXCEL_MAX_COL);//XFE1048577
        $sheet->getWriter()->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
        $sheet->getWriter()->write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                xmlns:r="'.$this->urlSchemaFormat.'/relationships">'
        );
        $sheet->getWriter()->write('<sheetPr filterMode="false">');
        $sheet->getWriter()->write('<pageSetUpPr fitToPage="false"/>');
        $sheet->getWriter()->write('</sheetPr>');
        $sheet->setMaxCellTagStart($sheet->getWriter()->fTell());
        $sheet->getWriter()->write('<dimension ref="A1:'.$maxCell.'"/>');
        $sheet->setMaxCellTagEnd($sheet->getWriter()->fTell());
        $sheet->getWriter()->write('<sheetViews>');
        $sheet->getWriter()->write(
            '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" 
            showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" 
            tabSelected="'.$selectedTab.'" topLeftCell="A1" view="normal" windowProtection="false" 
            workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">'
        );
        $sheet->getWriter()->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        $sheet->getWriter()->write('</sheetView>');
        $sheet->getWriter()->write('</sheetViews>');
        $sheet->getWriter()->write('<cols>');
        $sheet->getWriter()->write('<col collapsed="false" hidden="false" max="1025" min="1" style="0" width="11.5"/>');
        $sheet->getWriter()->write('</cols>');
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
        if (preg_match("/[H]{1,2}:[M]{1,2}/", $cellFormat)) {
            return 'datetime';
        }
        if (preg_match("/[M]{1,2}:[S]{1,2}/", $cellFormat)) {
            return 'datetime';
        }
        if (preg_match("/[YY]{2,4}/", $cellFormat)) {
            return 'date';
        }
        if (preg_match("/[D]{1,2}/", $cellFormat)) {
            return 'date';
        }
        if (preg_match("/[M]{1,2}/", $cellFormat)) {
            return 'date';
        }
        if (preg_match("/$/", $cellFormat)) {
            return 'currency';
        }
        if (preg_match("/%/", $cellFormat)) {
            return 'percent';
        }
        if (preg_match("/0/", $cellFormat)) {
            return 'numeric';
        }

        return 'string';
    }

    /**
     * @todo  check escaping
     *
     * @param $cellFormat
     *
     * @return string
     */
    private function escapeCellFormat($cellFormat)
    {
        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($cellFormat); $i < $ix; $i++) {
            $c = $cellFormat[$i];
            if ($ignoreUntil == '' && $c == '[') {
                $ignoreUntil = ']';
            } else {
                if ($ignoreUntil == '' && $c == '"') {
                    $ignoreUntil = '"';
                } else {
                    if ($ignoreUntil == $c) {
                        $ignoreUntil = '';
                    }
                }
            }
            if ($ignoreUntil == '' &&
                ($c == ' ' || $c == '-' || $c == '(' || $c == ')') &&
                ($i == 0 || $cellFormat[$i - 1] != '_')
            ) {
                $escaped .= "\\".$c;
            } else {
                $escaped .= $c;
            }
        }

        return $escaped;
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
        switch ($cellFormat) {
            case 'string':
                $cellFormat = 'GENERAL';
                break;
            case 'integer':
                $cellFormat = '0';
                break;
            case 'date':
                $cellFormat = 'YYYY-MM-DD';
                break;
            case 'datetime':
                $cellFormat = 'YYYY-MM-DD HH:MM:SS';
                break;
            case 'dollar':
                $cellFormat = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
                break;
            case 'money':
                $cellFormat = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
                break;
            case 'euro':
                $cellFormat = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
                break;
            case 'NN':
                $cellFormat = 'DDD';
                break;
            case 'NNN':
                $cellFormat = 'DDDD';
                break;
            case 'NNNN':
                $cellFormat = 'DDDD", "';
                break;
        }

        $cellFormat = strtoupper($cellFormat);
        $position = array_search($cellFormat, $this->cellFormats, $strict = true);
        if ($position === false) {
            $position = count($this->cellFormats);
            $this->cellFormats[] = $this->escapeCellFormat($cellFormat);
            $this->cellTypes[] = $this->determineCellType($cellFormat);
        }

        return $position;
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
            $sheet->setColumns(array_fill($from = 0, $until = count($row), '0'));
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
        if (empty($sheetName) || $this->sheets[$sheetName]->getFinalized()) {
            return;
        }
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $sheet->getWriter()->write('</sheetData>');
        $mergeCells = $sheet->getMergeCells();
        if (!empty($mergeCells)) {
            $sheet->getWriter()->write('<mergeCells>');
            foreach ($mergeCells as $range) {
                $sheet->getWriter()->write('<mergeCell ref="'.$range.'"/>');
            }
            $sheet->getWriter()->write('</mergeCells>');
        }
        $sheet->getWriter()->write(
            '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false"
                verticalCentered="false"/>'
        );
        $sheet->getWriter()->write(
            '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>'
        );
        $sheet->getWriter()->write(
            '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" 
                fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" 
                paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>'
        );
        $sheet->getWriter()->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->getWriter()->write(
            '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>'
        );
        $sheet->getWriter()->write(
            '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>'
        );
        $sheet->getWriter()->write('</headerFooter>');
        $sheet->getWriter()->write('</worksheet>');
        $maxCell = self::xlsCell($sheet->getRowCount() - 1, count($sheet->getColumns()) - 1);
        $maxCellTag = '<dimension ref="A1:'.$maxCell.'"/>';
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
        $startCell = self::xlsCell($startCellRow, $startCellColumn);
        $endCell = self::xlsCell($endCellRow, $endCellColumn);
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
     * @param $rowNumber
     * @param $columnNumber
     * @param $value
     * @param $cellIndex
     */
    protected function writeCell(
        Writer $file,
        $rowNumber,
        $columnNumber,
        $value,
        $cellIndex
    ) {
        $cellType = $this->cellTypes[$cellIndex];
        $cellName = self::xlsCell($rowNumber, $columnNumber);
        if (!is_scalar($value) || $value === '') {
            $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'"/>');
        } elseif (is_string($value) && $value{0} == '=') {
            $file->write(
                sprintf('<c r="%s" s="%s" t="s"><f>%s</f></c>', $cellName, $cellIndex, self::xmlspecialchars($value))
            );
        } elseif ($cellType == 'date') {
            $file->write(
                sprintf('<c r="%s" s="%s" t="n"><v>%s</v></c>', $cellName, $cellIndex, self::convertDateTime($value))
            );
        } elseif ($cellType == 'datetime') {
            $file->write(
                '<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.self::convertDateTime($value).'</v></c>'
            );
        } elseif ($cellType == 'currency' || $cellType == 'percent' || $cellType == 'numeric') {
            $file->write(
                '<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.self::xmlspecialchars($value).'</v></c>'
            );
        } else {
            if (!is_string($value)) {
                $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.($value * 1).'</v></c>');
            } else {
                if ($value{0} != '0' && $value{0} != '+' && filter_var(
                    $value,
                    FILTER_VALIDATE_INT,
                    ['options' => ['max_range' => self::EXCEL_MAX_RANGE]]
                )) {
                    $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.($value * 1).'</v></c>');
                } else {
                    $file->write(
                        '<c r="'.$cellName.'" s="'.$cellIndex.'" t="s"><v>'.self::xmlspecialchars(
                            $this->setSharedString($value)
                        ).'</v></c>'
                    );
                }
            }
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
     * @param int $rowNumber
     * @param int $columnNumber
     *
     * @return string Cell label/coordinates (A1, C3, AA42)
     */
    public static function xlsCell($rowNumber, $columnNumber)
    {
        $n = $columnNumber;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41).$r;
        }

        return $r.($rowNumber + 1);
    }

    /**
     * @param $string
     */
    public static function log($string)
    {
        file_put_contents(
            "php://stderr",
            date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n"
        );
    }

    /**
     * @link https://msdn.microsoft.com/ru-RU/library/aa365247%28VS.85%29.aspx
     *
     * @param string $filename
     *
     * @return mixed
     */
    public static function checkFilename($filename)
    {
        $invalidCharacter = array_merge(
            array_map('chr', range(0, 31)),
            ['<', '>', '?', '"', ':', '|', '\\', '/', '*', '&']
        );

        return str_replace($invalidCharacter, '', $filename);
    }

    /**
     * @param $val
     *
     * @return mixed
     */
    public static function xmlspecialchars($val)
    {
        return str_replace("'", "&#39;", htmlspecialchars($val));
    }

    /**
     * @param string $dateInput
     *
     * @return int
     */
    public static function convertDateTime($dateInput)
    {
        // Time expressed as fraction of 24h hours in seconds
        $seconds = 0;
        $year = $month = $day = 0;
        $dateTime = $dateInput;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $dateTime, $matches)) {
            list($year, $month, $day) = $matches;
        }
        if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $dateTime, $matches)) {
            list($hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }
        //using 1900 as epoch, not 1904, ignoring 1904 special case
        // Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31') {
            return $seconds;
        }    // Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00') {
            return $seconds;
        }    // Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29') {
            return 60 + $seconds;
        }
        // Excel false leapday
        /*
         We calculate the date by calculating the number of days since the epoch
         and adjust for the number of leap days. We calculate the number of leap
         days by normalising the year in relation to the epoch. Thus the year 2000
         becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        */
        $epoch  = 1900;
        $offset = 0;
        $norm   = 300;
        $range  = $year - $epoch;
        // Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
        $mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
        // Some boundary checks
        if ($year < $epoch || $year > 9999) {
            return 0;
        }
        if ($month < 1 || $month > 12) {
            return 0;
        }
        if ($day < 1 || $day > $mdays[$month - 1]) {
            return 0;
        }
        // Accumulate the number of days since the epoch.
        // Add days for current month
        $days = $day;
        // Add days for past months
        $days += array_sum(array_slice($mdays, 0, $month - 1));
        // Add days for past years
        $days += $range * 365;
        // Add leapdays
        $days += intval(($range) / 4);
        // Subtract 100 year leapdays
        $days -= intval(($range + $offset) / 100);
        // Add 400 year leapdays
        $days += intval(($range + $offset + $norm) / 400);
        // Already counted above
        $days -= $leap;
        // Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }
}
