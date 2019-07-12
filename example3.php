<?php

require __DIR__ . '/vendor/autoload.php';

use Zhaqq\Xlsx\Writer\Builder;

date_default_timezone_set('PRC');
$start = microtime(true);
times($start);

try {
    $writer   = new Builder();
    $fileName = __DIR__ . '/data/xlsx_writer' . date('Ymd-His') . '.xlsx';
    $writer->buildHeader('sheet_name_1', [
        'title'   => 'number',
        'content' => 'number',
        'weight'  => 'number',
    ]);

    foreach (rows() as $row) {
        $writer->writeSheetRow($row[0], $row[1]);
    }
    times($start);

    $writer->writeToFile($fileName);
    times($start);

} catch (\Exception $exception) {
    var_dump($exception->getMessage());
}

function rows($n = 20)
{
    for ($i = 0; $i < $n; $i++) {
        yield ['sheet_name_1', [
            $i,
            $i,
            '=A{n}+B{n}',
        ]];
    }
}

function times($start, $object = 'XlsxWriter')
{
    echo $object, PHP_EOL;
    echo microtime(true) - $start, PHP_EOL;
    echo '#', floor((memory_get_peak_usage(true)) / 1024 / 1024), "MB", PHP_EOL;
    echo '#', floor((memory_get_usage(true)) / 1024 / 1024), "MB", PHP_EOL, PHP_EOL;
}