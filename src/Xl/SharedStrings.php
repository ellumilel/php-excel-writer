<?php
namespace Ellumilel\Xl;

/**
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class Relationships
 * @package Ellumilel\Xl
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class SharedStrings
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    /** @var int */
    private $sharedStringsCount = 0;
    /** @var int */
    private $sharedStringsLength = 0;
    /** @var array */
    private $sharedStrings = [];

    /**
     * SharedStrings constructor.
     *
     * @param int $sharedStringsCount
     * @param array $sharedStrings
     */
    public function __construct($sharedStringsCount = 0, $sharedStrings = [])
    {
        $this->sharedStringsCount = $sharedStringsCount;
        $this->sharedStrings = $sharedStrings;
        $this->sharedStringsLength = count($sharedStrings);
    }

    /**
     * @return string
     */
    public function buildSharedStringsXML()
    {
        $ssXml = '';
        $ssXml .= $this->xml;
        $ssXml .= $this->getSst();

        return $ssXml;
    }

    /**
     * @return string
     */
    private function getSst()
    {
        $sst = '<sst count="'.
            $this->sharedStringsCount.'" uniqueCount="'.
            $this->sharedStringsLength.'" xmlns="'.
            $this->urlSchemaFormat.'">'
        ;

        foreach ($this->sharedStrings as $s => $item) {
            $sst .= '<si><t>'.str_replace("'", "&#39;", htmlspecialchars($s)).'</t></si>';
        }

        $sst .= '</sst>';

        return $sst;
    }
}
