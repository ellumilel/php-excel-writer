<?php
namespace Ellumilel\DocProps;

/**
 * @todo more Properties
 * @link https://msdn.microsoft.com/en-us/library/bb264572(v=office.12).aspx
 *
 * Class App
 * @package Ellumilel\DocProps
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class App
{
    /** @var string */
    private $xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    /** @var string */
    private $urlSchemaFormat = 'http://schemas.openxmlformats.org/officeDocument/2006';
    /** @var int */
    private $totalTime = 0;

    /**
     * App constructor.
     *
     * @param int $totalTime
     */
    public function __construct($totalTime = 0)
    {
        $this->totalTime = $totalTime;
    }

    /**
     * @return string
     */
    public function buildAppXML()
    {
        $appXml = '';
        $appXml .= $this->xml;
        $appXml .= $this->getAppProperties();

        return $appXml;
    }

    /**
     * @return string
     */
    private function getAppProperties()
    {
        $properties = '<Properties xmlns="'.
            $this->urlSchemaFormat.'/extended-properties" xmlns:vt="'.
            $this->urlSchemaFormat.'/docPropsVTypes">';
        $properties .= $this->getTotalTime();
        $properties .= '</Properties>';

        return $properties;
    }

    /**
     * @return string
     */
    private function getTotalTime()
    {
        return '<TotalTime>'.$this->totalTime.'</TotalTime>';
    }
}
