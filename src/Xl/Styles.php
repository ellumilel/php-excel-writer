<?php
namespace Ellumilel\Xl;

/**
 * @todo work with all styles
 *
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class Styles
 * @package Ellumilel\Xl
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class Styles
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlOpenXmlFormat = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    /** @var array */
    private $cellFormats;

    /**
     * @return string
     */
    public function buildStylesXML()
    {
        $xml = '';
        $xml .= $this->xml;
        $xml .= '<styleSheet xmlns="'.$this->urlOpenXmlFormat.'">';
        $xml .= '<numFmts count="'.count($this->cellFormats).'">';
        foreach ($this->cellFormats as $i => $v) {
            $xml .= '<numFmt numFmtId="'.(164 + $i).'" formatCode="'.
                str_replace("'", "&#39;", htmlspecialchars($v)).
                '" />';
        }
        $xml .= '</numFmts>';
        $xml .= $this->getFonts();
        $xml .= $this->getFills();
        $xml .= $this->getStyleXfs();
        $xml .= $this->getXfs();
        $xml .= $this->getCellStyles();
        $xml .= '</styleSheet>';

        return $xml;
    }

    /**
     * @return string
     */
    private function getFonts()
    {
        $fonts = '<fonts count="4">';
        $fonts .= '<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>';
        $fonts .= '<font><name val="Arial"/><family val="0"/><sz val="10"/></font>';
        $fonts .= '<font><name val="Arial"/><family val="0"/><sz val="10"/></font>';
        $fonts .= '<font><name val="Arial"/><family val="0"/><sz val="10"/></font>';
        $fonts .= '</fonts>';

        return $fonts;
    }

    /**
     * @return string
     */
    private function getFills()
    {
        $fills = '<fills count="2"><fill><patternFill patternType="none"/></fill>';
        $fills .= '<fill><patternFill patternType="gray125"/></fill></fills>';
        $fills .= '<borders count="1">';
        $fills .= '<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>';
        $fills .= '</borders>';

        return $fills;
    }

    /**
     * @return string
     */
    private function getStyleXfs()
    {
        $xfs = '<cellStyleXfs count="20">';
        $xfs .= '<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">';
        $xfs .= '<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>';
        $xfs .= '<protection hidden="false" locked="true"/>';
        $xfs .= '</xf>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>';
        $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>';
        $xfs .= '</cellStyleXfs>';

        return $xfs;
    }

    /**
     * @return string
     */
    private function getXfs()
    {
        $xfs = '<cellXfs count="'.count($this->cellFormats).'">';
        foreach ($this->cellFormats as $i => $v) {
            $xfs .= '<xf applyAlignment="false" applyBorder="false" applyFont="false" 
            applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="'.(164 + $i).'" xfId="0"/>';
        }
        $xfs .= '</cellXfs>';

        return $xfs;
    }

    /**
     * @return string
     */
    private function getCellStyles()
    {
        $style = '<cellStyles count="6">';
        $style .= '<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>';
        $style .= '<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>';
        $style .= '<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>';
        $style .= '<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>';
        $style .= '<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>';
        $style .= '<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>';
        $style .= '</cellStyles>';

        return $style;
    }

    /**
     * @param array $cellFormats
     *
     * @return $this
     */
    public function setCellFormats(array $cellFormats)
    {
        $this->cellFormats = $cellFormats;

        return $this;
    }
}
