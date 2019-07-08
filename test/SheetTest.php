<?php

declare(strict_types=1);

use Zhaqq\Xlsx\Writer\Sheet;

/**
 * Class SheetTest
 */
class SheetTest extends \PHPUnit\Framework\TestCase
{

    public function testFile()
    {
        $tempDir = __DIR__ . '/../data/';
        $filename = tempnam($tempDir, "xlsx_writer_");
        $sheet = new Sheet(
            [
                'filename' => $filename,
                'sheetname' => 'sheet_1',
                'xmlname' => 'sheet_1',
                'row_count' => 0,
                'columns' => [],
                'merge_cells' => [],
                'max_cell_tag_start' => 0,
                'max_cell_tag_end' => 0,
                'auto_filter' => false,
                'freeze_rows' => false,
                'freeze_columns' => false,
                'finalized' => false,
            ], 'xlsx'
        );

        $sheet->fileWriter->ftell();
        $this->assertSame(0, $sheet->fileWriter->ftell());
        $string = '';
        for($i=0;$i< 10;$i++) {
            $string .= 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
        }
        $sheet->fileWriter->write($string);
        $this->assertSame(620, $sheet->fileWriter->ftell());
    }
}
