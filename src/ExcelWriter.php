<?php
namespace xlsxWriter;

/**
 * Class ExcelWriter
 * @package xlsxWriter
 */
class ExcelWriter
{
    /**
     * @link http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     */
    const EXCEL_MAX_ROW = 1048576;
    const EXCEL_MAX_COL = 16384;

    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    /** @var string */
    protected $author ='Unknown Author';
    /** @var array */
    protected $sheets = [];
    /** @var array */
    protected $sharedStrings = [];//unique set
    /** @var int */
    protected $shared_string_count = 0;//count of non-unique references to the unique set
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
                @unlink($tempFile);
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
            self::finalizeSheet($sheetName);//making sure all footers have been written
        }
        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename); //if the zip already exists, remove it
            } else {
                self::log("Error in ".__CLASS__."::".__FUNCTION__.", file is not writeable.");

                return;
            }
        }
        $zip = new \ZipArchive();
        if (empty($this->sheets)) {
            self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined.");

            return;
        }
        if (!$zip->open($filename, \ZipArchive::CREATE)) {
            self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip.");

            return;
        }
        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", self::buildAppXML());
        $zip->addFromString("docProps/core.xml", self::buildCoreXML());
        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", self::buildRelationshipsXML());
        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            /** @var Sheet $sheet */
            $zip->addFile($sheet->getFilename(), "xl/worksheets/".$sheet->getXmlName());
        }
        if (!empty($this->sharedStrings)) {
            $zip->addFile(
                $this->writeSharedStringsXML(),
                "xl/sharedStrings.xml"
            );  //$zip->addFromString("xl/sharedStrings.xml",     self::buildSharedStringsXML() );
        }
        $zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
        $zip->addFile(
            $this->writeStylesXML(),
            "xl/styles.xml"
        );  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString("[Content_Types].xml", self::buildContentTypesXML());
        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelXML());
        $zip->close();
    }

    /**
     * @param string $sheetName
     */
    protected function initializeSheet($sheetName)
    {
        //if already initialized
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
        /*$this->sheets[$sheetName] = (object)[
            'filename' => $sheetFilename,
            'sheetname' => $sheetName,
            'xmlname' => $sheetXmlName,
            'row_count' => 0,
            'file_writer' => new Writer($sheetFilename),
            'columns' => [],
            'merge_cells' => [],
            'max_cell_tag_start' => 0,
            'max_cell_tag_end' => 0,
            'finalized' => false,
        ];*/
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $selectedTab = count($this->sheets) == 1 ? 'true' : 'false';//only first sheet is selected
        $maxCell = ExcelWriter::xlsCell(self::EXCEL_MAX_ROW, self::EXCEL_MAX_COL);//XFE1048577
        $sheet->getWriter()->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
        $sheet->getWriter()->write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                xmlns:r="'.$this->urlSchemaFormat.'">'
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
        //return str_replace( array(" ","-", "(", ")"), array("\ ","\-", "\(", "\)"), $cell_format);//TODO, needs more escaping
    }

    /**
     * @param $cellFormat
     *
     * @return int|mixed
     */
    private function addCellFormat($cellFormat)
    {
        //for backwards compatibility, to handle older versions
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
        self::initializeSheet($sheetName);
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $sheet->setColumns([]);
        foreach ($headerTypes as $v) {
            $sheet->setColumn($this->addCellFormat($v));
        }

        if (!$suppressRow) {
            $header_row = array_keys($headerTypes);
            $sheet->getWriter()->write(
                '<row collapsed="false" customFormat="false" 
                customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="'.(1).'">'
            );
            foreach ($header_row as $k => $v) {
                $this->writeCell($sheet->getWriter(), 0, $k, $v, $cell_format_index = '0');
            }
            $sheet->getWriter()->write('</row>');
            $sheet->increaseRowCount();
        }
        $this->currentSheet = $sheetName;
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
        self::initializeSheet($sheetName);
        /** @var Sheet $sheet */
        $sheet = &$this->sheets[$sheetName];
        $columns = $sheet->getColumns();
        if (empty($columns)) {
            $sheet->setColumns(array_fill($from = 0, $until = count($row), '0'));//'0'=>'string'
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
        $padding_length = $sheet->getMaxCellTagEnd() - $sheet->getMaxCellTagStart() - strlen($maxCellTag);
        $sheet->getWriter()->fSeek($sheet->getMaxCellTagStart());
        $sheet->getWriter()->write($maxCellTag.str_repeat(" ", $padding_length));
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
        self::initializeSheet($sheetName);
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
        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'"/>');
        } elseif (is_string($value) && $value{0} == '=') {
            $file->write(
                '<c r="'.$cellName.'" s="'.$cellIndex.'" t="s"><f>'.self::xmlspecialchars($value).'</f></c>'
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
            );//int,float,currency
        } else {
            if (!is_string($value)) {
                $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.($value * 1).'</v></c>');
            } else {
                if ($value{0} != '0' && $value{0} != '+' && filter_var(
                    $value,
                    FILTER_VALIDATE_INT,
                    ['options' => ['max_range' => 2147483647]]
                )) {
                    $file->write('<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.($value * 1).'</v></c>');
                } else { //implied: ($cell_format=='string')
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
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        $file->write('<numFmts count="'.count($this->cellFormats).'">');
        foreach ($this->cellFormats as $i => $v) {
            $file->write('<numFmt numFmtId="'.(164 + $i).'" formatCode="'.self::xmlspecialchars($v).'" />');
        }
        //$file->write(		'<numFmt formatCode="GENERAL" numFmtId="164"/>');
        //$file->write(		'<numFmt formatCode="[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00" numFmtId="165"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD\ HH:MM:SS" numFmtId="166"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD" numFmtId="167"/>');
        $file->write('</numFmts>');
        $file->write('<fonts count="4">');
        $file->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('</fonts>');
        $file->write('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
        $file->write('<borders count="1"><border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border></borders>');
        $file->write('<cellStyleXfs count="20">');
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $file->write('</cellStyleXfs>');
        $file->write('<cellXfs count="'.count($this->cellFormats).'">');
        foreach ($this->cellFormats as $i => $v) {
            $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" 
            applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="'.(164+$i).'" xfId="0"/>');
        }
        $file->write('</cellXfs>');
        //$file->write(	'<cellXfs count="4">');
        //$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>');
        //$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>');
        //$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"/>');
        //$file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="167" xfId="0"/>');
        //$file->write(	'</cellXfs>');
        $file->write('<cellStyles count="6">');
        $file->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $file->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $file->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $file->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $file->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $file->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $file->write('</cellStyles>');
        $file->write('</styleSheet>');
        $file->close();
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
        $this->shared_string_count++;//non-unique count

        return $stringValue;
    }

    /**
     * @return string
     */
    protected function writeSharedStringsXML()
    {
        $temporaryFilename = $this->tempFilename();
        $file = new Writer($temporaryFilename, $fd_flags = 'w', $check_utf8 = true);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
        $file->write(
            '<sst count="'.($this->shared_string_count).'" uniqueCount="'.count(
                $this->sharedStrings
            ).'" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        );
        foreach ($this->sharedStrings as $s => $c) {
            $file->write('<si><t>'.self::xmlspecialchars($s).'</t></si>');
        }
        $file->write('</sst>');
        $file->close();

        return $temporaryFilename;
    }

    /**
     * @return string
     */
    protected function buildAppXML()
    {
        $app_xml="";
        $app_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $app_xml.='<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" 
                    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                    <TotalTime>0</TotalTime></Properties>'
        ;
        return $app_xml;
    }

    /**
     * @return string
     */
    protected function buildCoreXML()
    {
        $core_xml = "";
        $core_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $core_xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        //$date_time = '2014-10-25T15:54:37.00Z';
        $core_xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';
        $core_xml .= '<dc:creator>'.self::xmlspecialchars($this->author).'</dc:creator>';
        $core_xml .= '<cp:revision>0</cp:revision>';
        $core_xml .= '</cp:coreProperties>';

        return $core_xml;
    }

    /**
     * @return string
     */
    protected function buildRelationshipsXML()
    {
        $relXml = "";
        $relXml .= '<?xml version="1.0" encoding="UTF-8"?>'."\n";
        $relXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relXml .= '<Relationship Id="rId1" Type="'.$this->urlSchemaFormat.'/officeDocument" 
        Target="xl/workbook.xml"/>';
        $relXml .= '<Relationship Id="rId2" Type="'.$this->urlSchemaFormat.'/metadata/core-properties" 
        Target="docProps/core.xml"/>';
        $relXml .= '<Relationship Id="rId3" Type="'.$this->urlSchemaFormat.'/extended-properties" 
        Target="docProps/app.xml"/>';
        $relXml .= "\n";
        $relXml .= '</Relationships>';

        return $relXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookXML()
    {
        $i = 0;
        $workbookXml = "";
        $workbookXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
        $workbookXml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="'.$this->urlSchemaFormat.'">';
        $workbookXml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" 
                showObjects="all" date1904="false"/><workbookProtection/>';
        $workbookXml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" 
                showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" 
                xWindow="0" yWindow="0"/></bookViews>';
        $workbookXml .= '<sheets>';
        foreach ($this->sheets as $sheet_name => $sheet) {
            /** @var Sheet $sheet */
            $workbookXml .= '<sheet name="'.self::xmlspecialchars($sheet->getSheetName()).'" 
            sheetId="'.($i + 1).'" state="visible" r:id="rId'.($i + 2).'"/>';
            $i++;
        }
        $workbookXml .= '</sheets>';
        $workbookXml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';

        return $workbookXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookRelXML()
    {
        $i = 0;
        $relXml = '';
        $relXml .= '<?xml version="1.0" encoding="UTF-8"?>'."\n";
        $relXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relXml .= '<Relationship Id="rId1" Type="'.$this->urlSchemaFormat.'/styles" Target="styles.xml"/>';
        foreach ($this->sheets as $sheet_name => $sheet) {
            /** @var Sheet $sheet */
            $relXml .= '<Relationship Id="rId'.($i + 2).'" 
            Type="'.$this->urlSchemaFormat.'/worksheet" Target="worksheets/'.($sheet->getXmlName()).'"/>';
            $i++;
        }
        if (!empty($this->sharedStrings)) {
            $relXml .= '<Relationship Id="rId'.(count($this->sheets) + 2).'" 
            Type="'.$this->urlSchemaFormat.'/sharedStrings" Target="sharedStrings.xml"/>';
        }
        $relXml .= "\n";
        $relXml .= '</Relationships>';

        return $relXml;
    }

    /**
     * @return string
     */
    protected function buildContentTypesXML()
    {
        $contentTypeXml = "";
        $contentTypeXml .= '<?xml version="1.0" encoding="UTF-8"?>'."\n";
        $contentTypeXml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $contentTypeXml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $contentTypeXml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($this->sheets as $sheet_name => $sheet) {
            /** @var Sheet $sheet */
            $contentTypeXml .= '<Override PartName="/xl/worksheets/'.($sheet->getXmlName()).'" 
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        if (!empty($this->sharedStrings)) {
            $contentTypeXml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        }
        $contentTypeXml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $contentTypeXml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $contentTypeXml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $contentTypeXml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $contentTypeXml .= "\n";
        $contentTypeXml .= '</Types>';

        return $contentTypeXml;
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
    public static function convertDateTime($dateInput) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        $days = 0;    # Number of days since epoch
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;
        $hour = $min = $sec = 0;
        $date_time = $dateInput;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }
        //using 1900 as epoch, not 1904, ignoring 1904 special case
        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31') {
            return $seconds;
        }    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00') {
            return $seconds;
        }    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29') {
            return 60 + $seconds;
        }    # Excel false leapday
        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch  = 1900;
        $offset = 0;
        $norm   = 300;
        $range  = $year - $epoch;
        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
        $mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );
        # Some boundary checks
        if ($year < $epoch || $year > 9999) {
            return 0;
        }
        if ($month < 1 || $month > 12) {
            return 0;
        }
        if ($day < 1 || $day > $mdays[$month - 1]) {
            return 0;
        }
        # Accumulate the number of days since the epoch.
        $days = $day; # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1)); # Add days for past months
        $days += $range * 365; # Add days for past years
        $days += intval(($range) / 4); # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400); # Add 400 year leapdays
        $days -= $leap; # Already counted above
        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }
}
