<?php

namespace Zhaqq\Xlsx;

use Zhaqq\Xlsx\Writer\Sheet;
use Zhaqq\Xlsx\Writer\WriterBufferInterface;
use Zhaqq\Xlsx\Writer\XlsxBuilder;
use Zhaqq\Xlsx\Writer\XlsxWriterBuffer;

/**
 * @see     http://www.ecma-international.org/publications/standards/Ecma-376.htm
 * @see     http://officeopenxml.com/SSstyles.php
 * @see     http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
 *
 * Class XlsxWriter
 * @package Home\Logic\CommodityManagementLogic\Excel
 */
class XlsxWriter
{

    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;

    /**
     * @var
     */
    protected $title;
    /**
     * @var
     */
    protected $subject;
    /**
     * @var
     */
    protected $author;
    /**
     * @var
     */
    protected $company;
    /**
     * @var
     */
    protected $description;
    /**
     * @var array
     */
    protected $keywords = [];
    /**
     * 默认Sheet Writer
     *
     * @var string
     */
    protected $writer = 'xlsx';
    /**
     * @var
     */
    protected $currentSheet;
    /**
     * @var Sheet[]
     */
    protected $sheets = [];
    /**
     * 缓存文件列表
     *
     * @var array
     */
    protected $tempFiles = [];
    /**
     * 风格设置
     *
     * @var array
     */
    protected $cellStyles = [];
    /**
     * @var array
     */
    protected $numberFormats = [];
    /**
     * @var string
     */
    protected $tempDir;

    protected $sheetName;
    /**
     * @var Sheet
     */
    protected $sheet;

    /**
     * XlsxWriter constructor.
     */
    public function __construct()
    {
        if (!ini_get('date.timezone')) {
            date_default_timezone_set('PRC');
        }
        $this->addCellStyle($numberFormat = 'GENERAL', $styleString = null);
        $this->addCellStyle($numberFormat = 'GENERAL', $styleString = null);
        $this->addCellStyle($numberFormat = 'GENERAL', $styleString = null);
        $this->addCellStyle($numberFormat = 'GENERAL', $styleString = null);
    }

    /**
     * 写到标准输出
     */
    public function writeToStdOut()
    {
        $tempFile = $this->tempFilename();
        self::writeToFile($tempFile);
        readfile($tempFile);
    }

    /**
     * 写至文本并返回字符串
     *
     * @return false|string
     */
    public function writeToString()
    {
        $tempFile = $this->tempFilename();
        self::writeToFile($tempFile);
        $string = file_get_contents($tempFile);

        return $string;
    }

    /**
     * 写入sheet 不带header头 需先执行 $this->writeSheetHeader 初始化头部
     *
     * @param \Generator|array $rows
     *
     * @example $rows = [
     *          'sheet_name' => 'sheet_name_1',
     *          'row' => [
     *               'title'    => 'title1',
     *               'content' => 'content1'
     *      ]
     * ];
     */
    public function addSheetRows($rows)
    {
        foreach ($rows as $row) {
            !isset($row['options']) && $row['options'] = null;
            $this->writeSheetRow($row['sheet_name'], $row['row'], $row['options']);
        }
    }

    /**
     * 写入sheet 带header头 如headers为空需先执行 $this->writeSheetHeader 初始化头部
     *
     * @param \Generator|array $rows
     * @param array $headers
     *
     * @example $rows = [
     *          'sheet_name' => 'sheet_name_1',
     *          'row' => [
     *               'title'    => 'title1',
     *               'content' => 'content1'
     *          ]
     * ];
     *          $headers => [
     *          'sheet_name' => 'sheet_name_1',
     *          'types' => [
     *               'title'    => 'string',   // 标明类型
     *               'content' => 'string'
     *          ]
     *
     * ];
     */
    public function addSheetRowsWithHeaders($rows, $headers = [])
    {
        if (!empty($headers) && is_array($headers)) {
            foreach ($headers as $header) {
                !isset($header['options']) && $header['options'] = null;
                $this->writeSheetHeader($header['sheet_name'], $header['types'], $header['options'] = null);
            }
        }
        foreach ($rows as $row) {
            !isset($row['options']) && $row['options'] = null;
            $this->writeSheetRow($row['sheet_name'], $row['row'], $row['options']);
        }
    }

    /**
     * 写入文件
     *
     * @param $filename
     */
    public function writeToFile($filename)
    {
        foreach ($this->sheets as $sheetName => $sheet) {
            self::finalizeSheet($sheetName);//making sure all footers have been written
        }

        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename); //if the zip already exists, remove it
            } else {
                throw new XlsxException("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
            }
        }
        $zip = new \ZipArchive();
        if (empty($this->sheets)) {
            throw new XlsxException("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", no worksheets defined.");
        }
        if (!$zip->open($filename, \ZipArchive::CREATE)) {
            throw new XlsxException("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
        }
        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", XlsxBuilder::buildAppXML());
        $zip->addFromString("docProps/core.xml", XlsxBuilder::buildCoreXML());
        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", XlsxBuilder::buildRelationshipsXML());
        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            $zip->addFile($sheet->filename, "xl/worksheets/" . $sheet->xmlname);
        }
        $zip->addFromString("xl/workbook.xml", XlsxBuilder::buildWorkbookXML($this->sheets));
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");  //$zip->addFromString("xl/styles.xml", self::buildStylesXML() );
        $zip->addFromString("[Content_Types].xml", XlsxBuilder::buildContentTypesXML($this->sheets));
        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML());

        $zip->close();
    }

    protected function initSheet(string $sheetName, array $colOptions = [], array $headerTypes = [])
    {
        if ($this->currentSheet == $sheetName || isset($this->sheets[$sheetName])) {
            return;
        }
        $colWidths = isset($colOptions['widths']) ? (array)$colOptions['widths'] : [];
        $this->createSheet($sheetName, $colOptions);
        $sheet = $this->sheets[$sheetName];
        $sheet->initContent($colWidths, $this->isTabSelected());
        if (!empty($headerTypes)) {
            $sheet->columns[] = $this->initColumnsTypes($headerTypes);
            $headerRow = array_keys($headerTypes);
            $writer = $sheet->getFileWriter();
            $writer->write(
                '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . 1 . '">'
            );
            foreach ($headerRow as $c => $v) {
                $cellStyleIdx = empty($style) ?
                    $sheet->columns[$c]['default_cell_style'] :
                    $this->addCellStyle('GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style));
                $sheet->writeCell(0, $c, $v, 'n_string', $cellStyleIdx);
            }
            $writer->write('</row>');
            $sheet->rowCount++;
        }
    }

    protected function initColumnsTypes($headerTypes)
    {
        foreach ($headerTypes as $v) {
            $numberFormat = Support::numberFormatStandardized($v);
            $cellStyleIdx = $this->addCellStyle($numberFormat, $styleString = null);
            $columns[] = [
                'number_format' => $numberFormat,      //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => Support::determineNumberFormatType($numberFormat), //contains friendly format like 'datetime'
                'default_cell_style' => $cellStyleIdx,
            ];
        }

        return $columns ?? [];
    }

    /**
     * 是否第一个sheet
     *
     * @return bool
     */
    public function isTabSelected()
    {
        return count($this->sheets) === 1;
    }


    /**
     * @param string $sheetName
     * @param array $colOptions
     */
    protected function createSheet(string $sheetName, array $colOptions = [])
    {
        $sheetFilename = $this->tempFilename();
        $sheetXmlName = 'sheet' . (count($this->sheets) + 1) . ".xml";
        $autoFilter = isset($colOptions['auto_filter']) ? intval($colOptions['auto_filter']) : false;
        $freezeRows = isset($colOptions['freeze_rows']) ? intval($colOptions['freeze_rows']) : false;
        $freezeColumns = isset($colOptions['freeze_columns']) ? intval($colOptions['freeze_columns']) : false;

        $this->sheets[$sheetName] = new Sheet(
            [
                'filename' => $sheetFilename,
                'sheetname' => $sheetName,
                'xmlname' => $sheetXmlName,
                'row_count' => 0,
                'columns' => [],
                'merge_cells' => [],
                'max_cell_tag_start' => 0,
                'max_cell_tag_end' => 0,
                'auto_filter' => $autoFilter,
                'freeze_rows' => $freezeRows,
                'freeze_columns' => $freezeColumns,
                'finalized' => false,
            ], 'xlsx'
        );

        $this->sheet = $this->sheets[$sheetName];
    }

    /**
     * @param $numberFormat
     * @param $cellStyleString
     *
     * @return false|int|string
     */
    private function addCellStyle($numberFormat, $cellStyleString)
    {
        $numberFormatIdx = self::add2listGetIndex($this->numberFormats, $numberFormat);
        $lookupString = $numberFormatIdx . ";" . $cellStyleString;

        return self::add2listGetIndex($this->cellStyles, $lookupString);
    }

    private function initializeColumnTypes($headerTypes)
    {
        $column_types = array();
        foreach ($headerTypes as $v) {
            $numberFormat = self::numberFormatStandardized($v);
            $numberFormat_type = self::determineNumberFormatType($numberFormat);
            $cellStyleIdx = $this->addCellStyle($numberFormat, $styleString = null);
            $column_types[] = [
                'number_format' => $numberFormat,      //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => $numberFormat_type, //contains friendly format like 'datetime'
                'default_cell_style' => $cellStyleIdx,
            ];
        }
        return $column_types;
    }

    /**
     * @param       $sheetName
     * @param array $headerTypes ['标题' => 'string', 'content' => 'string', 'cost' => 'number']
     * @param null $colOptions
     */
    public function writeSheetHeader($sheetName, array $headerTypes, $colOptions = null)
    {
        if (empty($sheetName) || empty($headerTypes) || !empty($this->sheets[$sheetName]))
            return;

        $suppress_row = isset($colOptions['suppress_row']) ? intval($colOptions['suppress_row']) : false;
        if (is_bool($colOptions)) {
            throw new XlsxException("passing $suppress_row=false|true to writeSheetHeader() is deprecated");
        }

        $style = &$colOptions;

        $this->initSheet($sheetName, $colOptions, $headerTypes);
        $sheet = &$this->sheets[$sheetName];
        $sheet->columns = $this->initializeColumnTypes($headerTypes);
        if (!$suppress_row) {
            $header_row = array_keys($headerTypes);

            $sheet->fileWriter->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
            foreach ($header_row as $c => $v) {
                $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle('GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style));
                $this->writeCell($sheet->fileWriter, 0, $c, $v, $numberFormat_type = 'n_string', $cell_style_idx);
            }
            $sheet->fileWriter->write('</row>');
            $sheet->rowCount++;
        }
        $this->currentSheet = $sheetName;
    }

    /**
     * @param       $sheetName
     * @param array $row
     * @param null $rowOptions
     */
    public function writeSheetRow($sheetName, array $row, $rowOptions = null)
    {
        if (empty($sheetName))
            return;

        $this->initSheet($sheetName);
        /* @var $sheet Sheet */
        $sheet = $this->sheets[$sheetName];
        if (count($sheet->columns) < count($row)) {
            $default_column_types = $this->initColumnsTypes(array_fill($from = 0, $until = count($row), 'GENERAL'));//will map to n_auto
            $sheet->columns = array_merge((array)$sheet->columns, $default_column_types);
        }

        if (!empty($rowOptions)) {
            $ht = isset($rowOptions['height']) ? floatval($rowOptions['height']) : 12.1;
            $customHt = isset($rowOptions['height']) ? true : false;
            $hidden = isset($rowOptions['hidden']) ? (bool)($rowOptions['hidden']) : false;
            $collapsed = isset($rowOptions['collapsed']) ? (bool)($rowOptions['collapsed']) : false;
            $sheet->fileWriter->write('<row collapsed="' . ($collapsed) . '" customFormat="false" customHeight="' . ($customHt) . '" hidden="' . ($hidden) . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($sheet->rowCount + 1) . '">');
        } else {
            $sheet->fileWriter->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->rowCount + 1) . '">');
        }

        $style = &$rowOptions;
        $c = 0;
        foreach ($row as $v) {

            $numberFormat = $sheet->columns[$c]['number_format'];
            $numberFormatType = $sheet->columns[$c]['number_format_type'];
            $cellStyleIdx = empty($style) ? $sheet->columns[$c]['default_cell_style'] :
                $this->addCellStyle($numberFormat, json_encode(isset($style[0]) ? $style[$c] : $style));
            $this->writeCell($sheet->fileWriter, $sheet->rowCount, $c, $v, $numberFormatType, $cellStyleIdx);
            $c++;
        }
        $sheet->fileWriter->write('</row>');
        $sheet->rowCount++;
        $this->currentSheet = $sheetName;
    }

    /**
     * @param string $sheetName
     *
     * @return int
     */
    public function countSheetRows($sheetName = '')
    {
        $sheetName = $sheetName ?: $this->currentSheet;

        return array_key_exists($sheetName, $this->sheets) ? $this->sheets[$sheetName]->row_count : 0;
    }

    /**
     * @param $sheetName
     */
    protected function finalizeSheet($sheetName)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->finalized)
            return;

        /* @var $sheet Sheet */
        $sheet = &$this->sheets[$sheetName];

        $sheet->fileWriter->write('</sheetData>');

        if (!empty($sheet->merge_cells)) {
            $sheet->fileWriter->write('<mergeCells>');
            foreach ($sheet->merge_cells as $range) {
                $sheet->fileWriter->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->fileWriter->write('</mergeCells>');
        }

        $max_cell = $this->xlsCell($sheet->rowCount - 1, count($sheet->columns) - 1);

        if ($sheet->autoFilter) {
            $sheet->fileWriter->write('<autoFilter ref="A1:' . $max_cell . '"/>');
        }

        $sheet->fileWriter->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $sheet->fileWriter->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $sheet->fileWriter->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $sheet->fileWriter->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->fileWriter->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $sheet->fileWriter->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $sheet->fileWriter->write('</headerFooter>');
        $sheet->fileWriter->write('</worksheet>');

        $max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
        $padding_length = $sheet->maxCellTagEnd - $sheet->maxCellTagStart - strlen($max_cell_tag);
        $sheet->fileWriter->fseek($sheet->maxCellTagStart);
        $sheet->fileWriter->write($max_cell_tag . str_repeat(" ", $padding_length));
        $sheet->fileWriter->close();
        $sheet->finalized = true;
    }

    /**
     * @param $sheetName
     * @param $startCellRow
     * @param $startCellColumn
     * @param $endCellRow
     * @param $endCellColumn
     */
    public function markMergedCell($sheetName, $startCellRow, $startCellColumn, $endCellRow, $endCellColumn)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->finalized)
            return;

        $this->initSheet($sheetName);
        $sheet = &$this->sheets[$sheetName];

        $startCell = $this->xlsCell($startCellRow, $startCellColumn);
        $endCell = $this->xlsCell($endCellRow, $endCellColumn);
        $sheet->mergeCells[] = $startCell . ":" . $endCell;
    }

    /**
     * @param array $data
     * @param string $sheetName
     * @param array $headerTypes
     */
    public function writeSheet(array $data, $sheetName = '', array $headerTypes = array())
    {
        $sheetName = empty($sheetName) ? 'Sheet1' : $sheetName;
        $data = empty($data) ? array(array('')) : $data;
        if (!empty($headerTypes)) {
            $this->writeSheetHeader($sheetName, $headerTypes);
        }
        foreach ($data as $i => $row) {
            $this->writeSheetRow($sheetName, $row);
        }
        $this->finalizeSheet($sheetName);
    }

    /**
     * @param WriterBufferInterface $file
     * @param                       $rowNumber
     * @param                       $columnNumber
     * @param                       $value
     * @param                       $numFormatType
     * @param                       $cellStyleIdx
     */
    protected function writeCell(WriterBufferInterface &$file, $rowNumber, $columnNumber, $value, $numFormatType, $cellStyleIdx)
    {
        $cell_name = self::xlsCell($rowNumber, $columnNumber);

        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '"/>');
        } elseif (is_string($value) && $value{0} == '=') {
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="s"><f>' . self::xmlSpecialChars($value) . '</f></c>');
        } elseif ($numFormatType == 'n_date') {
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="n"><v>' . intval(self::convertDateTime($value)) . '</v></c>');
        } elseif ($numFormatType == 'n_datetime') {
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::convertDateTime($value) . '</v></c>');
        } elseif ($numFormatType == 'n_numeric') {
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
        } elseif ($numFormatType == 'n_string') {
            $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t>' . self::xmlSpecialChars($value) . '</t></is></c>');
        } elseif ($numFormatType == 'n_auto' || 1) { //auto-detect unknown column types
            if (!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
                $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
            } else { //implied: ($cell_format=='string')
                $file->write('<c r="' . $cell_name . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t>' . self::xmlSpecialChars($value) . '</t></is></c>');
            }
        }
    }

    /**
     * @return array
     */
    protected function styleFontIndexes()
    {
        static $border_allowed = array('left', 'right', 'top', 'bottom');
        static $border_style_allowed = array('thin', 'medium', 'thick', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot');
        static $horizontal_allowed = array('general', 'left', 'right', 'justify', 'center');
        static $vertical_allowed = array('bottom', 'center', 'distributed', 'top');
        $default_font = array('size' => '10', 'name' => 'Arial', 'family' => '2');
        $fills = array('', '');//2 placeholders for static xml later
        $fonts = array('', '', '', '');//4 placeholders for static xml later
        $borders = array('');//1 placeholder for static xml later
        $style_indexes = array();
        foreach ($this->cellStyles as $i => $cellStyleString) {
            $semi_colon_pos = strpos($cellStyleString, ";");
            $numberFormatIdx = substr($cellStyleString, 0, $semi_colon_pos);
            $style_json_string = substr($cellStyleString, $semi_colon_pos + 1);
            $style = @json_decode($style_json_string, $as_assoc = true);

            $style_indexes[$i] = array('num_fmt_idx' => $numberFormatIdx);//initialize entry
            if (isset($style['border']) && is_string($style['border']))//border is a comma delimited str
            {
                $border_value['side'] = array_intersect(explode(",", $style['border']), $border_allowed);
                if (isset($style['border-style']) && in_array($style['border-style'], $border_style_allowed)) {
                    $border_value['style'] = $style['border-style'];
                }
                if (isset($style['border-color']) && is_string($style['border-color']) && $style['border-color'][0] == '#') {
                    $v = substr($style['border-color'], 1, 6);
                    $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                    $border_value['color'] = "FF" . strtoupper($v);
                }
                $style_indexes[$i]['border_idx'] = self::add2listGetIndex($borders, json_encode($border_value));
            }
            if (isset($style['fill']) && is_string($style['fill']) && $style['fill'][0] == '#') {
                $v = substr($style['fill'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $style_indexes[$i]['fill_idx'] = self::add2listGetIndex($fills, "FF" . strtoupper($v));
            }
            if (isset($style['halign']) && in_array($style['halign'], $horizontal_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['halign'] = $style['halign'];
            }
            if (isset($style['valign']) && in_array($style['valign'], $vertical_allowed)) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['valign'] = $style['valign'];
            }
            if (isset($style['wrap_text'])) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
            }

            $font = $default_font;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']);//floatval to allow "10.5" etc
            }
            if (isset($style['font']) && is_string($style['font'])) {
                if ($style['font'] == 'Comic Sans MS') {
                    $font['family'] = 4;
                }
                if ($style['font'] == 'Times New Roman') {
                    $font['family'] = 1;
                }
                if ($style['font'] == 'Courier New') {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                if (strpos($style['font-style'], 'bold') !== false) {
                    $font['bold'] = true;
                }
                if (strpos($style['font-style'], 'italic') !== false) {
                    $font['italic'] = true;
                }
                if (strpos($style['font-style'], 'strike') !== false) {
                    $font['strike'] = true;
                }
                if (strpos($style['font-style'], 'underline') !== false) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['color']) && is_string($style['color']) && $style['color'][0] == '#') {
                $v = substr($style['color'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $font['color'] = "FF" . strtoupper($v);
            }
            if ($font != $default_font) {
                $style_indexes[$i]['font_idx'] = self::add2listGetIndex($fonts, json_encode($font));
            }
        }
        return array('fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $style_indexes);
    }

    /**
     * @return bool|string
     */
    protected function writeStylesXML()
    {
        $r = self::styleFontIndexes();
        $fills = $r['fills'];
        $fonts = $r['fonts'];
        $borders = $r['borders'];
        $style_indexes = $r['styles'];

        $temporaryFilename = $this->tempFilename();
        $file = new XlsxWriterBuffer($temporaryFilename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        $file->write('<numFmts count="' . count($this->numberFormats) . '">');
        foreach ($this->numberFormats as $i => $v) {
            $file->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . self::xmlSpecialChars($v) . '" />');
        }
        $file->write('</numFmts>');
        $file->write('<fonts count="' . (count($fonts)) . '">');
        $file->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');

        foreach ($fonts as $font) {
            if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
                $f = json_decode($font, true);
                $file->write('<font>');
                $file->write('<name val="' . htmlspecialchars($f['name']) . '"/><charset val="1"/><family val="' . intval($f['family']) . '"/>');
                $file->write('<sz val="' . intval($f['size']) . '"/>');
                if (!empty($f['color'])) {
                    $file->write('<color rgb="' . strval($f['color']) . '"/>');
                }
                if (!empty($f['bold'])) {
                    $file->write('<b val="true"/>');
                }
                if (!empty($f['italic'])) {
                    $file->write('<i val="true"/>');
                }
                if (!empty($f['underline'])) {
                    $file->write('<u val="single"/>');
                }
                if (!empty($f['strike'])) {
                    $file->write('<strike val="true"/>');
                }
                $file->write('</font>');
            }
        }
        $file->write('</fonts>');

        $file->write('<fills count="' . (count($fills)) . '">');
        $file->write('<fill><patternFill patternType="none"/></fill>');
        $file->write('<fill><patternFill patternType="gray125"/></fill>');
        foreach ($fills as $fill) {
            if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
                $file->write('<fill><patternFill patternType="solid"><fgColor rgb="' . strval($fill) . '"/><bgColor indexed="64"/></patternFill></fill>');
            }
        }
        $file->write('</fills>');

        $file->write('<borders count="' . (count($borders)) . '">');
        $file->write('<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
        foreach ($borders as $border) {
            if (!empty($border)) { //fonts have an empty placeholder in the array to offset the static xml entry above
                $pieces = json_decode($border, true);
                $border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
                $border_color = !empty($pieces['color']) ? '<color rgb="' . strval($pieces['color']) . '"/>' : '';
                $file->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (array('left', 'right', 'top', 'bottom') as $side) {
                    $show_side = in_array($side, $pieces['side']) ? true : false;
                    $file->write($show_side ? "<$side style=\"$border_style\">$border_color</$side>" : "<$side/>");
                }
                $file->write('<diagonal/>');
                $file->write('</border>');
            }
        }
        $file->write('</borders>');

        $file->write('<cellStyleXfs count="20">');
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $file->write('</cellStyleXfs>');

        $file->write('<cellXfs count="' . (count($style_indexes)) . '">');
        foreach ($style_indexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
            $vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont = 'true';
            $borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
            $file->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $file->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $file->write('	<protection locked="true" hidden="false"/>');
            $file->write('</xf>');
        }
        $file->write('</cellXfs>');
        $file->write('<cellStyles count="6">');
        $file->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $file->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $file->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $file->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $file->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $file->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $file->write('</cellStyles>');
        $file->write('</styleSheet>');
        $file->close();

        return $temporaryFilename;
    }

    /**
     * @return string
     */
    protected function buildAppXML()
    {
        $app_xml = "";
        $app_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $app_xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' .
            ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $app_xml .= '<TotalTime>0</TotalTime>';
        $app_xml .= '<Company>' . self::xmlSpecialChars($this->company) . '</Company>';
        $app_xml .= '</Properties>';

        return $app_xml;
    }

    /**
     * @return string
     */
    protected function buildCoreXML()
    {
        $coreXml = "";
        $coreXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $coreXml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $coreXml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';//$date_time = '2014-10-25T15:54:37.00Z';
        $coreXml .= '<dc:title>' . self::xmlSpecialChars($this->title) . '</dc:title>';
        $coreXml .= '<dc:subject>' . self::xmlSpecialChars($this->subject) . '</dc:subject>';
        $coreXml .= '<dc:creator>' . self::xmlSpecialChars($this->author) . '</dc:creator>';
        if (!empty($this->keywords)) {
            $coreXml .= '<cp:keywords>' . self::xmlSpecialChars(implode(", ", (array)$this->keywords)) . '</cp:keywords>';
        }
        $coreXml .= '<dc:description>' . self::xmlSpecialChars($this->description) . '</dc:description>';
        $coreXml .= '<cp:revision>0</cp:revision>';
        $coreXml .= '</cp:coreProperties>';

        return $coreXml;
    }

    protected function buildRelationshipsXML()
    {
        $relsXml = "";
        $relsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $relsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $relsXml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $relsXml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $relsXml .= "\n";
        $relsXml .= '</Relationships>';

        return $relsXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookXML()
    {
        $i = 0;
        $workbookXml = "";
        $workbookXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbookXml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' .
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbookXml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbookXml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbookXml .= '<sheets>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $sheetname = self::sanitizeSheetname($sheet->sheetname);
            $workbookXml .= '<sheet name="' . self::xmlSpecialChars($sheetname) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
            $i++;
        }
        $workbookXml .= '</sheets>';
        $workbookXml .= '<definedNames>';
        foreach ($this->sheets as $sheetName => $sheet) {
            if ($sheet->autoFilter) {
                $sheetname = self::sanitizeSheetname($sheet->sheetname);
                $workbookXml .= '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' . self::xmlSpecialChars($sheetname) . '\'!$A$1:' . self::xlsCell($sheet->rowCount - 1, count($sheet->columns) - 1, true) . '</definedName>';
                $i++;
            }
        }
        $workbookXml .= '</definedNames>';
        $workbookXml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';

        return $workbookXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookRelsXML()
    {
        $i = 0;
        $wkbkrelsXml = "";
        $wkbkrelsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $wkbkrelsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $wkbkrelsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $wkbkrelsXml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet->xmlname) . '"/>';
            $i++;
        }
        $wkbkrelsXml .= "\n";
        $wkbkrelsXml .= '</Relationships>';
        return $wkbkrelsXml;
    }

    /**
     * @return string
     */
    protected function buildContentTypesXML()
    {
        $contentTypesXml = "";
        $contentTypesXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $contentTypesXml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $contentTypesXml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $contentTypesXml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlname) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $contentTypesXml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $contentTypesXml .= "\n";
        $contentTypesXml .= '</Types>';

        return $contentTypesXml;
    }

    /*
     * @param $rowNumber int, zero based
     * @param $columnNumber int, zero based
     * @param $absolute bool
     * @return Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     * */
    public static function xlsCell($rowNumber, $columnNumber, $absolute = false)
    {
        $n = $columnNumber;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        if ($absolute) {
            return '$' . $r . '$' . ($rowNumber + 1);
        }
        return $r . ($rowNumber + 1);
    }

    /**
     *
     * @param $filename
     *
     * @return mixed
     * @see http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
     *
     */
    public static function sanitizeFilename($filename)
    {
        $nonprinting = array_map('chr', range(0, 31));
        $invalid_chars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
        $all_invalids = array_merge($nonprinting, $invalid_chars);

        return str_replace($all_invalids, "", $filename);
    }

    /**
     * @param $sheetname
     *
     * @return string
     */
    public static function sanitizeSheetname($sheetname)
    {
        static $badchars = '\\/?*:[]';
        static $goodchars = '        ';
        $sheetname = strtr($sheetname, $badchars, $goodchars);
        $sheetname = mb_substr($sheetname, 0, 31);
        $sheetname = trim(trim(trim($sheetname), "'"));//trim before and after trimming single quotes

        return !empty($sheetname) ? $sheetname : 'Sheet' . ((rand() % 900) + 100);
    }

    /**
     * 检测xml一些特殊字符
     *
     * @param $val
     *
     * @return string
     */
    public static function xmlSpecialChars($val)
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodchars = "                              ";

        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badchars, $goodchars);//strtr appears to be faster than str_replace
    }

    /**
     * 返回数组第一个key
     *
     * @param array $arr
     *
     * @return int|string|null
     */
    public static function arrayFirstKey(array $arr)
    {
        reset($arr);

        return key($arr);
    }

    /**
     * 确定文件类型
     *
     * @param $numFormat
     *
     * @return string
     */
    private static function determineNumberFormatType($numFormat)
    {
        $numFormat = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $numFormat);
        if ($numFormat == 'GENERAL') return 'n_auto';
        if ($numFormat == '@') return 'n_string';
        if ($numFormat == '0') return 'n_numeric';
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat)) return 'n_datetime';
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat)) return 'n_datetime';
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat)) return 'n_date';
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat)) return 'n_date';
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat)) return 'n_date';
        if (preg_match('/$(?![^"]*+")/', $numFormat)) return 'n_numeric';
        if (preg_match('/%(?![^"]*+")/', $numFormat)) return 'n_numeric';
        if (preg_match('/0(?![^"]*+")/', $numFormat)) return 'n_numeric';
        return 'n_auto';

        switch ($numFormat) {
            case 'GENERAL':
                return 'n_auto';
            case '@':
                return 'n_string';
            case '0':
                return 'n_numeric';
            case preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat) != 0:
            case preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat) != 0:
                return 'n_datetime';
            case preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat) != 0:
            case preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat) != 0:
            case preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat) != 0:
                return 'n_date';
            case preg_match('/$(?![^"]*+")/', $numFormat) != 0:
            case preg_match('/%(?![^"]*+")/', $numFormat) != 0:
            case preg_match('/0(?![^"]*+")/', $numFormat) != 0:
                return 'n_numeric';
        }
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    private static function numberFormatStandardized($numFormat)
    {
        if ($numFormat == 'money') {
            $numFormat = 'dollar';
        }
        if ($numFormat == 'number') {
            $numFormat = 'integer';
        }
        switch ($numFormat) {
            case 'string':
                $numFormat = '@';
                break;
            case 'integer':
                $numFormat = '0';
                break;
            case 'date':
                $numFormat = 'YYYY-MM-DD';
                break;
            case 'datetime':
                $numFormat = 'YYYY-MM-DD HH:MM:SS';
                break;
            case 'price':
                $numFormat = '#,##0.00';
                break;
            case 'float3':
                $numFormat = '#,###0.000';
                break;
            case 'dollar':
                $numFormat = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
                break;
            case 'euro':
                $numFormat = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
                break;
        }
        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; $i++) {
            $c = $numFormat[$i];
            if ($ignoreUntil == '' && $c == '[')
                $ignoreUntil = ']';
            else if ($ignoreUntil == '' && $c == '"')
                $ignoreUntil = '"';
            else if ($ignoreUntil == $c)
                $ignoreUntil = '';
            if ($ignoreUntil == '' && ($c == ' ' || $c == '-' || $c == '(' || $c == ')') && ($i == 0 || $numFormat[$i - 1] != '_'))
                $escaped .= "\\" . $c;
            else
                $escaped .= $c;
        }

        return $escaped;
    }

    /**
     * @param $haystack
     * @param $needle
     *
     * @return false|int|string
     */
    public static function add2listGetIndex(&$haystack, $needle)
    {
        $existingIdx = array_search($needle, $haystack, $strict = true);
        if ($existingIdx === false) {
            $existingIdx = count($haystack);
            $haystack[] = $needle;
        }
        return $existingIdx;
    }

    /**
     * @param $dateInput
     *
     * @return float|int
     */
    public static function convertDateTime($dateInput)
    {
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;

        $date_time = $dateInput;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }
        unset($junk);

        //using 1900 as epoch, not 1904, ignoring 1904 special case
        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31') return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00') return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29') return 60 + $seconds;    # Excel false leapday
        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;
        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100))) ? 1 : 0;
        $mdays = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
        # Some boundary checks
        if ($year < $epoch || $year > 9999) return 0;
        if ($month < 1 || $month > 12) return 0;
        if ($day < 1 || $day > $mdays[$month - 1]) return 0;
        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += intval(($range) / 4);             # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above
        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }

    /**
     * @param string $writer
     *
     * @return $this
     */
    public function setWriter($writer = 'xlsx')
    {
        $this->writer = $writer;

        return $this;
    }

    /**
     * @param string $title
     *
     * @return $this
     */
    public function setTitle($title = '')
    {
        $this->title = $title;
        return $this;
    }

    /**
     * @param string $subject
     *
     * @return $this
     */
    public function setSubject($subject = '')
    {
        $this->subject = $subject;
        return $this;
    }

    /**
     * @param string $author
     *
     * @return $this
     */
    public function setAuthor($author = '')
    {
        $this->author = $author;
        return $this;
    }

    /**
     * @param string $company
     *
     * @return $this
     */
    public function setCompany($company = '')
    {
        $this->company = $company;
        return $this;
    }

    /**
     * @param string $keywords
     *
     * @return $this
     */
    public function setKeywords($keywords = '')
    {
        $this->keywords = $keywords;
        return $this;
    }

    /**
     * @param string $description
     *
     * @return $this
     */
    public function setDescription($description = '')
    {
        $this->description = $description;
        return $this;
    }

    /**
     * @param string $tempDir
     *
     * @return $this
     */
    public function setTempDir($tempDir = '')
    {
        $this->tempDir = $tempDir;
        return $this;
    }

    /**
     * @return bool|string
     */
    protected function tempFilename()
    {
        $tempDir = !empty($this->tempDir) ? $this->tempDir : sys_get_temp_dir();
        $filename = tempnam($tempDir, "xlsx_writer_");
        $this->tempFiles[] = $filename;

        return $filename;
    }

    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                @unlink($tempFile);
            }
        }
    }
}
