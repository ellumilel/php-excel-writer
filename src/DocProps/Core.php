<?php
namespace Ellumilel\DocProps;

/**
 * @link https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#Document_properties
 *
 * Class Core
 * @package Ellumilel\DocProps
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class Core
{
    /** @var string */
    private $author = 'Unknown Author';
    /** @var int */
    private $revision = 0;
    /** @var string */
    private $urlCp = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
    /** @var string */
    private $urlDc = 'http://purl.org/dc/elements/1.1/';
    /** @var string */
    private $urlDcType = 'http://purl.org/dc/dcmitype/';
    /** @var string */
    private $urlDcTerms = 'http://purl.org/dc/terms/';
    /** @var string */
    private $urlSchema = 'http://www.w3.org/2001/XMLSchema-instance';
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

    /**
     * Core constructor.
     *
     * @param integer $revision
     */
    public function __construct($revision = 0)
    {
        $this->revision = $revision;
    }

    /**
     * @return string
     */
    public function buildCoreXML()
    {
        $coreXml = '';
        $coreXml .= $this->xml;
        $coreXml .= $this->getCoreProperties();

        return $coreXml;
    }

    /**
     * @return string
     */
    private function getCoreProperties()
    {
        $properties = '<cp:coreProperties xmlns:cp="'.$this->urlCp.'" xmlns:dc="'.
            $this->urlDc.'" xmlns:dcmitype="'.$this->urlDcType.'" xmlns:dcterms="'.
            $this->urlDcTerms.'" xmlns:xsi="'.$this->urlSchema.'">';
        $properties .= $this->created();
        $properties .= $this->creator();
        $properties .= $this->revision();
        $properties .= '</cp:coreProperties>';

        return $properties;
    }

    /**
     * CreatedDate ex: '2016-08-30T15:52:19.00Z';
     * @return string
     */
    private function created()
    {
        return '<dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';
    }

    /**
     * @return string
     */
    private function creator()
    {
        return '<dc:creator>'.str_replace("'", "&#39;", htmlspecialchars($this->author)).'</dc:creator>';
    }

    /**
     * @return string
     */
    private function revision()
    {
        return '<cp:revision>'.$this->revision.'</cp:revision>';
    }

    /**
     * @param string $author
     *
     * @return $this
     */
    public function setAuthor($author)
    {
        if (!empty($author)) {
            $this->author = $author;
        }

        return $this;
    }
}
