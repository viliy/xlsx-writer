<?php

namespace Zhaqq\Xlsx\Writer;

use Zhaqq\Exception\XlsxException;

/**
 * Class XlsxWriterBuffer.
 *
 * @author  QiuXiaoYong on 2019-07-02 13:21
 */
class XlsxWriterBuffer implements WriterBufferInterface
{
    const BUFFER_SIZE = 8191;

    /**
     * 默认打开标识.
     */
    const F_OPEN_FLAGS_W = 'w';

    /**
     * 当前文件标识位.
     *
     * @var bool|resource|null
     */
    protected $fd = null;

    /**
     * 当前数据大小.
     *
     * @var string
     */
    protected $buffer = '';
    /**
     * 检验字符集.
     *
     * @var bool
     */
    protected $checkUtf8 = false;

    /**
     * XlsxWriterBuffer constructor.
     *
     * @param        $filename
     * @param string $fdFopenFlags
     * @param bool   $checkUtf8
     */
    public function __construct($filename, $fdFopenFlags = self::F_OPEN_FLAGS_W, $checkUtf8 = false)
    {
        $this->checkUtf8 = $checkUtf8;
        $this->fd        = fopen($filename, $fdFopenFlags);
        if (false === $this->fd) {
            throw new XlsxException("Unable to open $filename for writing.");
        }
    }

    /**
     * @param $string
     *
     * @author   QiuXiaoYong on 2019-07-02 14:09
     */
    public function write($string)
    {
        $this->buffer .= $string;
        if (isset($this->buffer[self::BUFFER_SIZE])) {
            $this->purge();
        }
    }

    /**
     * @author   QiuXiaoYong on 2019-07-02 14:09
     */
    protected function purge()
    {
        if ($this->fd) {
            if ($this->checkUtf8 && !self::isValidUTF8($this->buffer)) {
                $this->checkUtf8 = false;
                throw new XlsxException('Error, invalid UTF8 encoding detected.');
            }
            /** @scrutinizer ignore-type */
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    /**
     * @author   QiuXiaoYong on 2019-07-02 14:15
     */
    public function close()
    {
        $this->purge();
        if ($this->fd) {
            /** @scrutinizer ignore-type */
            fclose($this->fd);
            $this->fd = null;
        }
    }

    /**
     * @return bool|int
     *
     * @author   QiuXiaoYong on 2019-07-02 14:10
     */
    public function ftell()
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
     *
     * @author   QiuXiaoYong on 2019-07-02 14:16
     */
    public function fseek($pos)
    {
        if ($this->fd) {
            $this->purge();
            /** @scrutinizer ignore-type */
            return fseek($this->fd, $pos);
        }

        return -1;
    }

    /**
     * 关闭文件.
     */
    public function __destruct()
    {
        $this->close();
    }

    /**
     * @param $string
     *
     * @return bool
     *
     * @author   QiuXiaoYong on 2019-07-02 14:16
     */
    protected static function isValidUTF8($string)
    {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }

        return preg_match('//u', $string) ? true : false;
    }
}
