<?php
namespace Ellumilel\Rels;

use Ellumilel\Sheet;

/**
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class Relationships
 * @package Ellumilel\Rels
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class Relationships
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlSchema = 'http://schemas.openxmlformats.org/officeDocument/2006';
    /** @var string */
    private $urlRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
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
     * @return string
     */
    public function buildRelationshipsXML()
    {
        $relXml = '';
        $relXml .= $this->xml;
        $relXml .= '<Relationships xmlns="'.$this->urlRel.'">';
        $relXml .= '<Relationship Id="rId1" Type="'.$this->urlSchema.'/relationships/officeDocument"';
        $relXml .= ' Target="xl/workbook.xml"/>';
        $relXml .= '<Relationship Id="rId2" Type="'.$this->urlSchema.'/relationships/metadata/core-properties"';
        $relXml .= ' Target="docProps/core.xml"/>';
        $relXml .= '<Relationship Id="rId3" Type="'.$this->urlSchema.'/relationships/extended-properties"';
        $relXml .= ' Target="docProps/app.xml"/>'."\n";
        $relXml .= '</Relationships>';

        return $relXml;
    }

    /**
     * @return string
     */
    public function buildWorkbookRelationshipsXML()
    {
        $i = 0;
        $relXml = '';
        $relXml .= $this->xml;
        $relXml .= '<Relationships xmlns="'.$this->urlRel.'">';
        $relXml .= '<Relationship Id="rId1" Type="'.$this->urlSchema.'/relationships/styles" Target="styles.xml"/>';
        foreach ($this->sheets as $sheet) {
            /** @var Sheet $sheet */
            $relXml .= '<Relationship Id="rId'.($i + 2).'" 
            Type="'.$this->urlSchema.'/relationships/worksheet" Target="worksheets/'.($sheet->getXmlName()).'"/>';
            $i++;
        }
        if (!empty($this->sharedStringsRelations)) {
            $relXml .= '<Relationship Id="rId'.(count($this->sheets) + 2).'" 
            Type="'.$this->urlSchema.'/relationships/sharedStrings" Target="sharedStrings.xml"/>';
        }
        $relXml .= "\n";
        $relXml .= '</Relationships>';

        return $relXml;
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
