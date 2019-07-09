<?php

/**
 * Class Builder.
 */

namespace Zhaqq\Xlsx\Writer;

use Zhaqq\Exception\XlsxException;
use Zhaqq\Xlsx\Support;

/**
 * Class Builder.
 */
class Builder
{
    use Style;

    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;

    /**
     * @var Sheet[]
     */
    protected $sheets = [];

    /**
     * @var Sheet
     */
    protected $sheet;
    /**
     * @var
     */
    protected $sheetName;

    /**
     * @var string
     */
    protected $sheetWriter = 'Xlsx';

    /**
     * 缓存文件列表.
     *
     * @var array
     */
    protected $tempFiles = [];

    /**
     * @var
     */
    protected $title = '';
    /**
     * @var
     */
    protected $subject = '';
    /**
     * @var
     */
    protected $author = '';
    /**
     * @var
     */
    protected $company = '';
    /**
     * @var
     */
    protected $description = '';
    /**
     * @var array
     */
    protected $keywords = [];

    protected $tempDir;

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
     * 写入sheet 不带header头 需先执行 $this->writeSheetHeader 初始化头部.
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
     * 写入sheet 带header头 如headers为空需先执行 $this->writeSheetHeader 初始化头部.
     *
     * @param \Generator|array $rows
     * @param array            $headers
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
                $this->buildHeader($header['sheet_name'], $header['types'], $header['options'] = null);
            }
        }
        foreach ($rows as $row) {
            !isset($row['options']) && $row['options'] = null;
            $this->writeSheetRow($row['sheet_name'], $row['row'], $row['options']);
        }
    }

    /**
     * 写入文件.
     *
     * @param $filename
     */
    public function writeToFile($filename)
    {
        foreach ($this->sheets as $sheetName => $sheet) {
            $this->finalizeSheet($sheetName); //making sure all footers have been written
        }

        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename); //if the zip already exists, remove it
            } else {
                throw new XlsxException('Error in ' . __CLASS__ . '::' . __FUNCTION__ . ', file is not writeable.');
            }
        }
        $zip = new \ZipArchive();
        if (empty($this->sheets)) {
            throw new XlsxException('Error in ' . __CLASS__ . '::' . __FUNCTION__ . ', no worksheets defined.');
        }
        if (!$zip->open($filename, \ZipArchive::CREATE)) {
            throw new XlsxException('Error in ' . __CLASS__ . '::' . __FUNCTION__ . ', unable to create zip.');
        }
        $zip->addEmptyDir('docProps/');
        $zip->addFromString('docProps/app.xml', XlsxBuilder::buildAppXML($this->company));
        $zip->addFromString('docProps/core.xml', XlsxBuilder::buildCoreXML(
            $this->title,
            $this->subject,
            $this->author,
            $this->keywords,
            $this->description
        ));
        $zip->addEmptyDir('_rels/');
        $zip->addFromString('_rels/.rels', XlsxBuilder::buildRelationshipsXML());
        $zip->addEmptyDir('xl/worksheets/');
        foreach ($this->sheets as $sheet) {
            $zip->addFile($sheet->filename, 'xl/worksheets/' . $sheet->xmlname);
        }
        $zip->addFromString('xl/workbook.xml', XlsxBuilder::buildWorkbookXML($this->sheets));
        $zip->addFile($this->writeStylesXML(), 'xl/styles.xml');  //$zip->addFromString("xl/styles.xml", self::buildStylesXML() );
        $zip->addFromString('[Content_Types].xml', XlsxBuilder::buildContentTypesXML($this->sheets));
        $zip->addEmptyDir('xl/_rels/');
        $zip->addFromString('xl/_rels/workbook.xml.rels', XlsxBuilder::buildWorkbookRelsXML($this->sheets));

        $zip->close();
    }

    /**
     * @param       $sheetName
     * @param array $row
     * @param array $rowOptions
     */
    public function writeSheetRow($sheetName, array $row, $rowOptions = [])
    {
        if (empty($sheetName)) {
            return;
        }

        $this->initSheet($sheetName);
        $sheet = $this->getSheet();
        if (count($sheet->columns) < count($row)) {
            $defaultColumnTypes = $this->initColumnsTypes(array_fill($from = 0, count($row), 'GENERAL')); //will map to n_auto
            $sheet->columns     = array_merge((array)$sheet->columns, $defaultColumnTypes);
        }

        if (!empty($rowOptions)) {
            $ht        = isset($rowOptions['height']) ? floatval($rowOptions['height']) : 12.1;
            $customHt  = isset($rowOptions['height']) ? true : false;
            $hidden    = isset($rowOptions['hidden']) ? (bool)($rowOptions['hidden']) : false;
            $collapsed = isset($rowOptions['collapsed']) ? (bool)($rowOptions['collapsed']) : false;
            $sheet->fileWriter->write(
                '<row collapsed="' . ($collapsed) . '" customFormat="false" customHeight="' . ($customHt) .
                '" hidden="' . ($hidden) . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($sheet->rowCount + 1) . '">'
            );
        } else {
            $sheet->fileWriter->write(
                '<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' .
                ($sheet->rowCount + 1) . '">'
            );
        }

        $style = $rowOptions;
        $c     = 0;
        foreach ($row as $v) {
            $numberFormat     = $sheet->columns[$c]['number_format'];
            $numberFormatType = $sheet->columns[$c]['number_format_type'];
            $cellStyleIdx     = empty($style) ? $sheet->columns[$c]['default_cell_style'] :
                $this->addCellStyle($numberFormat, json_encode(isset($style[0]) ? $style[$c] : $style));
            $sheet->writeCell($sheet->rowCount, $c, $v, $numberFormatType, $cellStyleIdx);
            ++$c;
        }
        $sheet->fileWriter->write('</row>');
        ++$sheet->rowCount;

        $this->sheetName = $sheetName;
    }

    /**
     * @param $sheetName
     */
    protected function finalizeSheet($sheetName)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->finalized) {
            return;
        }

        $sheet = $this->sheets[$sheetName];
        $sheet->finallyContent();
    }

    /**
     * @return bool|string
     */
    protected function writeStylesXML()
    {
        $r            = $this->styleFontIndexes($this->cellStyles);
        $fills        = $r['fills'];
        $fonts        = $r['fonts'];
        $borders      = $r['borders'];
        $styleIndexes = $r['styles'];

        $temporaryFilename = $this->tempFilename();
        $file              = new XlsxWriterBuffer($temporaryFilename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        $file->write('<numFmts count="' . count($this->numberFormats) . '">');
        foreach ($this->numberFormats as $i => $v) {
            $file->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . Support::xmlSpecialChars($v) . '" />');
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
                $pieces       = json_decode($border, true);
                $border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
                $border_color = !empty($pieces['color']) ? '<color rgb="' . strval($pieces['color']) . '"/>' : '';
                $file->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (['left', 'right', 'top', 'bottom'] as $side) {
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

        $file->write('<cellXfs count="' . (count($styleIndexes)) . '">');
        foreach ($styleIndexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText       = !empty($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
            $vertAlignment  = isset($v['valign']) ? $v['valign'] : 'bottom';
            $applyBorder    = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont      = 'true';
            $borderIdx      = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx        = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx        = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
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
     * @param string $sheetName
     * @param array  $headerTypes
     * @param array  $colOptions
     */
    public function buildHeader(string $sheetName, array $headerTypes, $colOptions = [])
    {
        if (empty($sheetName) || empty($headerTypes) || $this->hasSheet($sheetName)) {
            return;
        }
        $this->initSheet($sheetName, $colOptions, $headerTypes);
        $this->sheetName = $sheetName;
    }

    /**
     * @param string $sheetName
     * @param array  $headerTypes
     * @param array  $colOptions
     */
    protected function initSheet(string $sheetName, array $colOptions = [], array $headerTypes = [])
    {
        if ($this->sheetName == $sheetName || $this->hasSheet($sheetName)) {
            return;
        }
        $style     = $colOptions;
        $colWidths = isset($colOptions['widths']) ? (array)$colOptions['widths'] : [];
        $this->createSheet($sheetName, $colOptions);
        $sheet = $this->getSheet();
        $sheet->initContent($colWidths, $this->isTabSelected());
        if (!empty($headerTypes)) {
            $sheet->columns = $this->initColumnsTypes($headerTypes);
            $headerRow      = array_keys($headerTypes);
            $writer         = $sheet->getFileWriter();
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
            ++$sheet->rowCount;
        }
    }

    protected function initColumnsTypes($headerTypes)
    {
        foreach ($headerTypes as $v) {
            $numberFormat = Support::numberFormatStandardized($v);
            $cellStyleIdx = $this->addCellStyle($numberFormat, $styleString = null);
            $columns[]    = [
                'number_format'      => $numberFormat,      //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => Support::determineNumberFormatType($numberFormat), //contains friendly format like 'datetime'
                'default_cell_style' => $cellStyleIdx,
            ];
        }

        return $columns ?? [];
    }

    /**
     * 是否第一个sheet.
     *
     * @return bool
     */
    public function isTabSelected()
    {
        return 1 === count($this->sheets);
    }

    /**
     * @param string|null $sheetName
     *
     * @return Sheet
     */
    public function getSheet(string $sheetName = null): Sheet
    {
        if ($sheetName) {
            $this->sheet = $this->sheets[$sheetName];
        }

        return $this->sheet;
    }

    /**
     * @param Sheet $sheet
     */
    public function setSheet(Sheet $sheet)
    {
        $this->sheet = $sheet;
    }

    /**
     * @param string $string
     */
    protected function sheetWriter(string $string)
    {
        $this->sheet->fileWriter->write($string);
    }

    /**
     * @param string $sheetName
     * @param array  $colOptions
     */
    protected function createSheet(string $sheetName, array $colOptions = [])
    {
        $sheetFilename = $this->tempFilename();
        $sheetXmlName  = 'sheet' . (count($this->sheets) + 1) . '.xml';
        $autoFilter    = isset($colOptions['auto_filter']) ? intval($colOptions['auto_filter']) : false;
        $freezeRows    = isset($colOptions['freeze_rows']) ? intval($colOptions['freeze_rows']) : false;
        $freezeColumns = isset($colOptions['freeze_columns']) ? intval($colOptions['freeze_columns']) : false;

        $this->sheets[$sheetName] = new Sheet(
            [
                'filename'           => $sheetFilename,
                'sheetname'          => $sheetName,
                'xmlname'            => $sheetXmlName,
                'row_count'          => 0,
                'columns'            => [],
                'merge_cells'        => [],
                'max_cell_tag_start' => 0,
                'max_cell_tag_end'   => 0,
                'auto_filter'        => $autoFilter,
                'freeze_rows'        => $freezeRows,
                'freeze_columns'     => $freezeColumns,
                'finalized'          => false,
            ],
            $this->sheetWriter
        );

        $this->sheet = $this->sheets[$sheetName];
    }

    /**
     * @param string $sheetName
     *
     * @return bool
     */
    protected function hasSheet(string $sheetName)
    {
        if (isset($this->sheets[$sheetName])) {
            $this->setSheet($this->sheets[$sheetName]);

            return true;
        }

        return false;
    }

    /**
     * @return Sheet[]
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * @param Sheet[] $sheets
     */
    public function setSheets($sheets)
    {
        $this->sheets = $sheets;
    }

    /**
     * @return bool|string
     */
    protected function tempFilename()
    {
        $tempDir           = !empty($this->tempDir) ? $this->tempDir : sys_get_temp_dir();
        $filename          = tempnam($tempDir, 'xlsx_writer_');
        $this->tempFiles[] = $filename;

        return $filename;
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
     * @param array $keywords
     *
     * @return $this
     */
    public function setKeywords(array $keywords = [])
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

    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                /** @scrutinizer ignore-unhandled */
                @unlink($tempFile);
            }
        }
    }
}
