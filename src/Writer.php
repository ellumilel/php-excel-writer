<?php
namespace Ellumilel;

/**
 * Class Writer
 * @package xlsxWriter
 */
class Writer
{
    /** @var null|resource */
    protected $fd = null;
    /** @var string */
    protected $buffer = '';
    /** @var int */
    protected $bufferSize;
    /** @var bool */
    protected $check_utf8 = false;

    /**
     * Writer constructor.
     *
     * @param string $filename
     * @param string $openFlags
     * @param int $bufferSize
     *
     * @throws \Exception
     */
    public function __construct($filename, $openFlags = 'w', $bufferSize = 8191)
    {
        $this->fd = fopen($filename, $openFlags);
        $this->bufferSize = $bufferSize;

        if ($this->fd === false) {
            throw new \Exception("Unable to open $filename for writing.");
        }
    }

    /**
     * @param $string
     */
    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[$this->bufferSize])) {
            $this->purge();
        }
    }

    /**
     * add to file
     */
    protected function purge()
    {
        if ($this->fd) {
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    /**
     * close writing
     */
    public function close()
    {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }

    /**
     * close after end
     */
    public function __destruct()
    {
        $this->close();
    }

    /**
     * @return int
     */
    public function fTell()
    {
        if ($this->fd) {
            $this->purge();

            return ftell($this->fd);
        }

        return -1;
    }

    /**
     * @param $pos
     *
     * @return int
     */
    public function fSeek($pos)
    {
        if ($this->fd) {
            $this->purge();

            return fseek($this->fd, $pos);
        }

        return -1;
    }
}
