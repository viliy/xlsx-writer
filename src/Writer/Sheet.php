<?php

namespace Zhaqq\Xlsx\Writer;

use Zhaqq\Xlsx\Support;

/**
 * @property mixed filename
 * @property mixed sheetname
 * @property mixed xmlname
 * @property mixed rowCount
 * @property mixed columns
 * @property mixed mergeCells
 * @property mixed maxCellTagStart
 * @property mixed maxCellTagEnd
 * @property mixed autoFilter
 * @property mixed freezeRows
 * @property mixed freezeColumns
 * @property mixed finalized
 */
class Sheet
{
    /**
     * @var WriterBufferInterface
     */
    public $fileWriter;

    /**
     * Sheet constructor.
     *
     * @param array       $config
     * @param string|null $fileWriter
     */
    public function __construct(array $config, $fileWriter = null)
    {
        if (empty($config)) {
            throw new \RuntimeException('sheet config must be array');
        }
        $this->filename = $config['filename'];
        $this->sheetname = $config['sheetname'];
        $this->xmlname = $config['xmlname'];
        $this->rowCount = $config['row_count'];
        $this->columns = $config['columns'];
        $this->mergeCells = $config['merge_cells'];
        $this->maxCellTagStart = $config['max_cell_tag_start'];
        $this->maxCellTagEnd = $config['max_cell_tag_end'];
        $this->autoFilter = $config['auto_filter'];
        $this->freezeRows = $config['freeze_rows'];
        $this->freezeColumns = $config['freeze_columns'];
        $this->finalized = $config['finalized'];
        $this->setFileWriter($fileWriter);
    }

    /**
     * @param $rowNumber
     * @param $columnNumber
     * @param $value
     * @param $numFormatType
     * @param $cellStyleIdx
     */
    public function writeCell($rowNumber, $columnNumber, $value, $numFormatType, $cellStyleIdx)
    {
        $cellName = Support::xlsCell($rowNumber, $columnNumber);
        $file = $this->getFileWriter();

        if (!is_scalar($value) || '' === $value) { //objects, array, empty
            $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'"/>');
        } elseif (is_string($value) && '=' == $value[0]) {
            // Support Formula
            $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="s"><f>'.
                str_replace('{n}', $rowNumber + 1, substr(Support::xmlSpecialChars($value), 1))
                .'</f></c>');
        } else {
            switch ($numFormatType) {
                case 'n_date':
                    $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="n"><v>'.
                        intval(Support::convertDateTime($value)).'</v></c>');

                    break;
                case 'n_datetime':
                    $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="n"><v>'.
                        Support::convertDateTime($value).'</v></c>');

                    break;
                case 'n_numeric':
                    $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="n"><v>'.
                        Support::xmlSpecialChars($value).'</v></c>'); //int,float,currency
                    break;
                case 'n_string':
                    $file->write('<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="inlineStr"><is><t>'.
                        Support::xmlSpecialChars($value).'</t></is></c>');

                    break;
                case 'n_auto':
                default: //auto-detect unknown column types
                    if (!is_string($value) || '0' == $value || ('0' != $value[0] && ctype_digit($value)) || preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)
                    ) { //int,float,currency
                        $file->write(
                            '<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="n"><v>'.
                            Support::xmlSpecialChars($value).'</v></c>'
                        );
                    } else { //implied: ($cell_format=='string')
                        $file->write(
                            '<c r="'.$cellName.'" s="'.$cellStyleIdx.'" t="inlineStr"><is><t>'.
                            Support::xmlSpecialChars($value).'</t></is></c>'
                        );
                    }

                    break;
            }
        }
    }

    /**
     * @param array $colWidths
     * @param bool  $isTabSelected
     */
    public function initContent(array $colWidths = [], $isTabSelected = false)
    {
        $writer = $this->getFileWriter();
        $tabSelected = $isTabSelected ? 'true' : 'false';
        $maxCell = Support::xlsCell(Builder::EXCEL_2007_MAX_ROW, Builder::EXCEL_2007_MAX_COL); //XFE1048577
        $writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
        $writer->write(
            <<<'EOF'
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr>
EOF
        );
        $this->maxCellTagStart = $this->fileWriter->ftell();
        $writer->write('<dimension ref="A1:'.$maxCell.'"/>');
        $this->maxCellTagEnd = $this->fileWriter->ftell();
        $writer->write('<sheetViews>');
        $writer->write(
            <<<EOF
<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="$tabSelected" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">'
EOF
        );
        if ($this->freezeRows && $this->freezeColumns) {
            $this->writeFreezeRowsAndColumns();
        } elseif ($this->freezeRows) {
            $this->writeFreezeRows();
        } elseif ($this->freezeColumns) {
            $this->writeFreezeColumns();
        } else { // not frozen
            $writer->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }
        $writer->write('</sheetView></sheetViews><cols>');
        $i = 0;
        if (!empty($colWidths)) {
            foreach ($colWidths as $colWidth) {
                $writer->write(
                    '<col collapsed="false" hidden="false" max="'.($i + 1).'" min="'.($i + 1).
                    '" style="0" customWidth="true" width="'.floatval($colWidth).'"/>'
                );
                ++$i;
            }
        }
        $writer->write(
            '<col collapsed="false" hidden="false" max="1024" min="'.($i + 1).
            '" style="0" customWidth="false" width="11.5"/>'.'</cols><sheetData>'
        );
    }

    public function finallyContent()
    {
        $this->fileWriter->write('</sheetData>');

        if (!empty($this->mergeCells)) {
            $this->fileWriter->write('<mergeCells>');
            foreach ($this->mergeCells as $range) {
                $this->fileWriter->write('<mergeCell ref="'.$range.'"/>');
            }
            $this->fileWriter->write('</mergeCells>');
        }

        $maxCell = Support::xlsCell($this->rowCount - 1, count($this->columns) - 1);

        if ($this->autoFilter) {
            $this->fileWriter->write('<autoFilter ref="A1:'.$maxCell.'"/>');
        }

        $this->fileWriter->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $this->fileWriter->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $this->fileWriter->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $this->fileWriter->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $this->fileWriter->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $this->fileWriter->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $this->fileWriter->write('</headerFooter>');
        $this->fileWriter->write('</worksheet>');

        $maxCellTag = '<dimension ref="A1:'.$maxCell.'"/>';
        $paddingLength = $this->maxCellTagEnd - $this->maxCellTagStart - strlen($maxCellTag);
        $this->fileWriter->fseek($this->maxCellTagStart);
        $this->fileWriter->write($maxCellTag.str_repeat(' ', $paddingLength));
        $this->fileWriter->close();
        $this->finalized = true;
    }

    protected function writeFreezeRowsAndColumns()
    {
        $writer = $this->getFileWriter();
        $writer->write(
            '<pane ySplit="'.$this->freezeRows.'" xSplit="'.$this->freezeColumns.
            '" topLeftCell="'.Support::xlsCell($this->freezeRows, $this->freezeColumns).
            '" activePane="bottomRight" state="frozen"/>'.'<selection activeCell="'.Support::xlsCell($this->freezeRows, 0).
            '" activeCellId="0" pane="topRight" sqref="'.Support::xlsCell($this->freezeRows, 0).'"/>'.
            '<selection activeCell="'.Support::xlsCell(0, $this->freezeColumns).
            '" activeCellId="0" pane="bottomLeft" sqref="'.Support::xlsCell(0, $this->freezeColumns).'"/>'.
            '<selection activeCell="'.Support::xlsCell($this->freezeRows, $this->freezeColumns).
            '" activeCellId="0" pane="bottomRight" sqref="'.Support::xlsCell($this->freezeRows, $this->freezeColumns).'"/>'
        );
    }

    protected function writeFreezeRows()
    {
        $writer = $this->getFileWriter();
        $writer->write(
            '<pane ySplit="'.$this->freezeRows.'" topLeftCell="'.
            Support::xlsCell($this->freezeRows, 0).'" activePane="bottomLeft" state="frozen"/>'.
            '<selection activeCell="'.Support::xlsCell($this->freezeRows, 0).
            '" activeCellId="0" pane="bottomLeft" sqref="'.Support::xlsCell($this->freezeRows, 0).'"/>'
        );
    }

    protected function writeFreezeColumns()
    {
        $writer = $this->getFileWriter();
        $writer->write(
            '<pane xSplit="'.$this->freezeColumns.'" topLeftCell="'.
            Support::xlsCell(0, $this->freezeColumns).'" activePane="topRight" state="frozen"/>'.
            '<selection activeCell="'.Support::xlsCell(0, $this->freezeColumns).
            '" activeCellId="0" pane="topRight" sqref="'.Support::xlsCell(0, $this->freezeColumns).'"/>'
        );
    }

    /**
     * @param string|null $fileWriter
     *
     * @return Sheet
     */
    public function setFileWriter($fileWriter = null)
    {
        switch ($fileWriter) {
            case 'xlsx':
            default:
                // 当前只实现 XlsxWriterBuffer 写入方式 后期可添加redis等方案 继承WriterBufferInterface即可
                $this->fileWriter = new XlsxWriterBuffer($this->filename);
        }

        return $this;
    }

    /**
     * @return WriterBufferInterface
     */
    public function getFileWriter()
    {
        return $this->fileWriter;
    }
}
