<?php
namespace Ellumilel;

/**
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class ContentTypes
 * @package Ellumilel
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class ContentTypes
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlContentTypes = 'http://schemas.openxmlformats.org/package/2006/content-types';
    /** @var bool */
    private $sharedStringsRelations = false;
    /** @var array */
    private $sheets = [];

    /**
     * Relationships constructor.
     *
     * @param bool $sharedStringsRelations
     */
    public function __construct($sharedStringsRelations = false)
    {
        $this->sharedStringsRelations = $sharedStringsRelations;
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

    /**
     * @return string
     */
    public function buildContentTypesXML()
    {
        $xml = '';
        $xml .= $this->xml;
        $xml .= '<Types xmlns="'.$this->urlContentTypes.'">';
        $xml .= '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $xml .= '<Default Extension="xml" ContentType="application/xml"/>';
        $xml .= '<Override PartName="/_rels/.rels"';
        $xml .= ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $xml .= '<Override PartName="/xl/_rels/workbook.xml.rels"';
        $xml .= ' ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';

        /** @var Sheet $sheet */
        foreach ($this->sheets as $sheet) {
            $xml .= '<Override PartName="/xl/worksheets/'.($sheet->getXmlName()).'"';
            $xml .= ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }

        if ($this->sharedStringsRelations) {
            $xml .= '<Override PartName="/xl/sharedStrings.xml"';
            $xml .= ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
        }

        $xml .= '<Override PartName="/xl/workbook.xml"';
        $xml .= ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $xml .= '<Override PartName="/xl/styles.xml"';
        $xml .= ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $xml .= '<Override PartName="/docProps/app.xml"';
        $xml .= ' ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $xml .= '<Override PartName="/docProps/core.xml"';
        $xml .= ' ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'."\n";
        $xml .= '</Types>';

        return $xml;
    }
}
