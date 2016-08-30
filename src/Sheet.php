<?php
namespace Ellumilel;

/**
 * Class Sheet
 * @package Ellumilel
 * @author Denis Tikhonov <ozy@mailserver.ru>
 */
class Sheet
{
    /** @var string */
    private $filename;
    /** @var string */
    private $sheetName;
    /** @var string */
    private $xmlName;
    /** @var int */
    private $rowCount = 0;
    /** @var Writer */
    private $writer;
    /** @var array */
    private $columns = [];
    /** @var array */
    private $mergeCells = [];
    /** @var int */
    private $maxCellTagStart = 0;
    /** @var int */
    private $maxCellTagEnd = 0;
    /** @var bool */
    private $finalized = false;

    /**
     * @return string
     */
    public function getFilename()
    {
        return $this->filename;
    }

    /**
     * @param string $filename
     *
     * @return $this
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;

        return $this;
    }

    /**
     * @return string
     */
    public function getSheetName()
    {
        return $this->sheetName;
    }

    /**
     * @param string $sheetName
     *
     * @return $this
     */
    public function setSheetName($sheetName)
    {
        $this->sheetName = $sheetName;

        return $this;
    }

    /**
     * @return string
     */
    public function getXmlName()
    {
        return $this->xmlName;
    }

    /**
     * @param string $xmlName
     *
     * @return $this
     */
    public function setXmlName($xmlName)
    {
        $this->xmlName = $xmlName;

        return $this;
    }

    /**
     * @return int
     */
    public function getRowCount()
    {
        return $this->rowCount;
    }

    /**
     * @param int $rowCount
     *
     * @return $this
     */
    public function setRowCount($rowCount)
    {
        $this->rowCount = $rowCount;

        return $this;
    }

    /**
     * @return int
     */
    public function increaseRowCount()
    {
        return $this->rowCount++;
    }

    /**
     * @return Writer
     */
    public function getWriter()
    {
        return $this->writer;
    }

    /**
     * @param Writer $writer
     *
     * @return $this
     */
    public function setWriter(Writer $writer)
    {
        $this->writer = $writer;

        return $this;
    }

    /**
     * @return array
     */
    public function getColumns()
    {
        return $this->columns;
    }

    /**
     * @param int $column
     *
     * @return $this
     */
    public function setColumn($column)
    {
        $this->columns[] = $column;

        return $this;
    }

    /**
     * @param array $columns
     *
     * @return $this
     */
    public function setColumns(array $columns)
    {
        $this->columns = $columns;

        return $this;
    }

    /**
     * @return array
     */
    public function getMergeCells()
    {
        return $this->mergeCells;
    }

    /**
     * @param string $mergeCells
     *
     * @return $this
     */
    public function setMergeCells($mergeCells)
    {
        $this->mergeCells[] = $mergeCells;

        return $this;
    }

    /**
     * @return int
     */
    public function getMaxCellTagStart()
    {
        return $this->maxCellTagStart;
    }

    /**
     * @param int $maxCellTagStart
     *
     * @return $this
     */
    public function setMaxCellTagStart($maxCellTagStart)
    {
        $this->maxCellTagStart = $maxCellTagStart;

        return $this;
    }

    /**
     * @return int
     */
    public function getMaxCellTagEnd()
    {
        return $this->maxCellTagEnd;
    }

    /**
     * @param int $maxCellTagEnd
     *
     * @return $this
     */
    public function setMaxCellTagEnd($maxCellTagEnd)
    {
        $this->maxCellTagEnd = $maxCellTagEnd;

        return $this;
    }

    /**
     * @return bool
     */
    public function getFinalized()
    {
        return $this->finalized;
    }

    /**
     * @param bool $finalized
     *
     * @return $this
     */
    public function setFinalized($finalized)
    {
        $this->finalized = $finalized;

        return $this;
    }
}
