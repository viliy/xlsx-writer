<?php

namespace Zhaqq\Xlsx\Writer;

/**
 * sheet文件写入类.
 *
 * Interface WriterBufferInterface
 *
 * @author  QiuXiaoYong on 2019-07-02 14:20
 */
interface WriterBufferInterface
{
    /**
     * 写入文件.
     *
     * @param $string
     *
     * @author   QiuXiaoYong on 2019-07-02 14:08
     */
    public function write($string);

    /**
     * 关闭文件写入.
     *
     *
     * @author   QiuXiaoYong on 2019-07-02 14:09
     */
    public function close();

    /**
     * 当前文件指针位置.
     *
     * @return bool|int
     *
     * @author   QiuXiaoYong on 2019-07-02 14:07
     */
    public function ftell();

    /**
     * 改变当前文件指针位置.
     *
     * @param $pos
     *
     * @return int
     *
     * @author   QiuXiaoYong on 2019-07-02 14:07
     */
    public function fseek($pos);
}
