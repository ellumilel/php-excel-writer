<?php
namespace Ellumilel\Xl;

use Ellumilel\Sheet;

/**
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class Workbook
 * @package Ellumilel\Xl
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class Workbook
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/officeDocument/2006';
    /** @var string */
    private $urlOpenXmlFormat = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    /** @var array */
    private $sheets = [];

    /**
     * @return string
     */
    public function buildWorkbookXML()
    {
        $i = 0;
        $xml = '';
        $xml .= $this->xml;
        $xml .= '<workbook xmlns="'.$this->urlOpenXmlFormat.'" xmlns:r="'.$this->urlSchemaFormat.'/relationships">';
        $xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false"';
        $xml .= ' showObjects="all" date1904="false"/><workbookProtection/>';
        $xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true"';
        $xml .= ' showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192"';
        $xml .= ' windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $xml .= '<sheets>';
        /** @var Sheet $sheet */
        foreach ($this->sheets as $sheet_name => $sheet) {
            $xml .= '<sheet name="'.str_replace("'", "&#39;", htmlspecialchars($sheet->getSheetName())).'"';
            $xml .= ' sheetId="'.($i + 1).'" state="visible" r:id="rId'.($i + 2).'"/>';
            $i++;
        }
        $xml .= '</sheets>';
        $xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';

        return $xml;
    }

    /**
     * @param array $sheets
     *
     * @return $this
     */
    public function setSheet(array $sheets)
    {
        $this->sheets = $sheets;

        return $this;
    }
}
