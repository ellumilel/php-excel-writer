<?php
namespace Ellumilel\Xl\Worksheets;

use Ellumilel\Helpers\ExcelHelper;

/**
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class SheetXml
 * @package Ellumilel\Xl\Worksheets
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class SheetXml
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlOpenXmlFormat = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/officeDocument/2006';

    /**
     * @return string
     */
    public function getXml()
    {
        return $this->xml;
    }

    /**
     * @return string
     */
    public function getSheetPr()
    {
        $sPr = '<sheetPr filterMode="false">';
        $sPr .= '<pageSetUpPr fitToPage="false"/>';
        $sPr .= '</sheetPr>';

        return $sPr;
    }

    /**
     * @return string
     */
    public function getWorksheet()
    {
        $wSheet = '<worksheet xmlns="'.$this->urlOpenXmlFormat.'" xmlns:r="'.$this->urlSchemaFormat.'/relationships">';

        return $wSheet;
    }

    /**
     * @param string $selectedTab
     *
     * @return string
     */
    public function getSheetViews($selectedTab)
    {
        $sViews = '<sheetViews>';
        $sViews .= '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false"';
        $sViews .= ' showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true"';
        $sViews .= ' tabSelected="'.$selectedTab.'" topLeftCell="A1" view="normal" windowProtection="false"';
        $sViews .= ' workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">';
        $sViews .= '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>';
        $sViews .= '</sheetView>';
        $sViews .= '</sheetViews>';

        return $sViews;
    }

    /**
     * @return string
     */
    public function getCools()
    {
        $sCols = '<cols>';
        $sCols .= '<col collapsed="false" hidden="false" max="1025" min="1" style="0" width="11.5"/>';
        $sCols .= '</cols>';

        return $sCols;
    }

    /**
     * @param string $maxCell
     *
     * @return string
     */
    public function getDimension($maxCell)
    {
        $sCols = '<dimension ref="A1:'.$maxCell.'"/>';

        return $sCols;
    }

    /**
     * @todo refactor
     *
     * @return string
     */
    public function getHeaderFooter()
    {
        $hf = '<headerFooter differentFirst="false" differentOddEven="false">';
        $hf .= '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>';
        $hf .= '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>';
        $hf .= '</headerFooter>';

        return $hf;
    }

    /**
     * @return string
     */
    public function getPageSetup()
    {
        $ps = '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1"';
        $ps .= ' fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver"';
        $ps .= ' paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>';

        return $ps;
    }

    /**
     * @return string
     */
    public function getPageMargins()
    {
        return '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>';
    }

    /**
     * @return string
     */
    public function getPrintOptions()
    {
        return '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false"
                verticalCentered="false"/>';
    }

    /**
     * @param array $mergeCells
     *
     * @return string
     */
    public function getMergeCells(array $mergeCells)
    {
        $mc = '<mergeCells>';
        foreach ($mergeCells as $range) {
            $mc .= '<mergeCell ref="'.$range.'"/>';
        }
        $mc .= '</mergeCells>';

        return $mc;
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $cellType
     * @param $value
     *
     * @return bool|string
     */
    public function getCell($cellName, $cellIndex, $cellType, $value)
    {
        if ($cellType == '@') {
            return false;
        }

        if (!is_scalar($value) || $value === '') {
            return '<c r="'.$cellName.'" s="'.$cellIndex.'"/>';
        }

        if (is_string($value) && $value{0} == '=') {
            return $this->getFormulaCell($cellName, $cellIndex, $value);
        }

        if ($cellType == 'date' || $cellType == 'datetime') {
            return $this->getDateCell($cellName, $cellIndex, $value);
        } elseif ($cellType == 'currency' || $cellType == 'percent' || $cellType == 'numeric') {
            return $this->getCurrencyCell($cellName, $cellIndex, $value);
        }

        return $this->checkIntCell($cellName, $cellIndex, $value);
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $value
     *
     * @return bool|string
     */
    private function checkIntCell($cellName, $cellIndex, $value)
    {
        if (!is_string($value)) {
            return $this->getIntCell($cellName, $cellIndex, $value);
        } else {
            if ($value{0} != '0' &&
                $value{0} != '+' &&
                filter_var(
                    $value,
                    FILTER_VALIDATE_INT,
                    ['options' => ['max_range' => ExcelHelper::EXCEL_MAX_RANGE]]
                )
            ) {
                return $this->getIntCell($cellName, $cellIndex, $value);
            } else {
                return false;
            }
        }
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $value
     *
     * @return string
     */
    private function getDateCell($cellName, $cellIndex, $value)
    {
        return sprintf(
            '<c r="%s" s="%s" t="n"><v>%s</v></c>',
            $cellName,
            $cellIndex,
            ExcelHelper::convertDateTime($value)
        );
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $value
     *
     * @return string
     */
    private function getCurrencyCell($cellName, $cellIndex, $value)
    {
        return sprintf(
            '<c r="%s" s="%s" t="n"><v>%s</v></c>',
            $cellName,
            $cellIndex,
            ExcelHelper::xmlspecialchars($value)
        );
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $value
     *
     * @return string
     */
    private function getIntCell($cellName, $cellIndex, $value)
    {
        return '<c r="'.$cellName.'" s="'.$cellIndex.'" t="n"><v>'.intval($value).'</v></c>';
    }

    /**
     * @param $cellName
     * @param $cellIndex
     * @param $value
     *
     * @return string
     */
    private function getFormulaCell($cellName, $cellIndex, $value)
    {
        return sprintf(
            '<c r="%s" s="%s" t="s"><f>%s</f></c>',
            $cellName,
            $cellIndex,
            ExcelHelper::xmlspecialchars($value)
        );
    }
}
