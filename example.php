<?php

require __DIR__ . '/vendor/autoload.php';

use Zhaqq\Xlsx\XlsxWriter;

date_default_timezone_set('PRC');
$start = microtime(true);
ini_set('memory_limit', '512M');

times($start);

try {
    $writer = new \Zhaqq\Xlsx\Writer\Builder();
//    $writer = new XlsxWriter();

    $fileName = __DIR__ . '/data/xlsx_writer' . date('Ymd-His') . '.xlsx';
    $writer->buildHeader('非服装', otherHead());
    $writer->buildHeader('服装', clothingHead());

    foreach (rows() as $row) {
        $writer->writeSheetRow($row[0], $row[1]);
    }
    times($start);

    $writer->writeToFile($fileName);

    times($start);

} catch (\Exception $exception) {

}


function rows()
{
    for ($i = 0; $i < 20; $i++) {
        if ($i % 2) {
            yield ['sheet_name_1', [
                'SKU' . $i,
                '尺码' . $i,
                '净重' . $i,
                '单价' . $i,
            ]];
        } else {
            yield ['sheet_name_2', [
                'SKU' . $i,
                '尺码' . $i,
                '净重' . $i,
                '单价' . $i,

            ]];
        }
    }
}

function rowsE()
{
    for ($i = 0; $i < 100; $i++) {
        if ($i % 2) {
            yield ['非服装', [
                'sku' => 'SKU' . $i,
                'skc' => 'SKC' . $i,
                'size' => '尺码' . $i,
                'real_weight' => '净重' . $i,
                'us_cost' => '单价' . $i,
            ]];
        } else {
            yield ['服装', [
                'sku' => 'SKU' . $i,
                'size' => '尺码' . $i,
                'real_weight' => '净重' . $i,
                'us_cost' => '单价' . $i,
            ]];
        }
    }
}

function times($start, $object = 'XlsxWriter')
{
    echo $object, PHP_EOL;
    echo microtime(true) - $start, PHP_EOL;
    echo '#', floor((memory_get_peak_usage(true)) / 1024 / 1024), "MB", PHP_EOL;
    echo '#', floor((memory_get_usage(true)) / 1024 / 1024), "MB", PHP_EOL, PHP_EOL;
}


function clothingHead()
{
    return [
        'SKU' => 'string',
        '尺码' => 'string',
        '净重' => 'string',
        '单价' => 'float3',
    ];
}

function otherHead()
{
    return [
        'SKU' => 'string',
        '尺码' => 'string',  // sku
        '净重' => 'string',  // sku
        '单价' => 'prices',  // sku
    ];
}

function word($i)
{
    return [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'Q', 'O', 'P', 'Q', 'R', 'S'
    ][$i];
}
