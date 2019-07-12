# xlsx - writer

[![Build Status](https://travis-ci.org/viliy/xlsx-writer.svg?branch=master)](https://travis-ci.org/viliy/xlsx-writer)
[![StyleCI](https://github.styleci.io/repos/195766403/shield?branch=master)](https://github.styleci.io/repos/195766403)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/?branch=master)
[![Code Coverage](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/coverage.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/?branch=master)
[![Build Status](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/build.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/build-status/master)
[![Code Intelligence Status](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/code-intelligence.svg?b=master)](https://scrutinizer-ci.com/code-intelligence)


* PHP >= 7.1
* see to [PHP_XLSXWriter](https://github.com/mk-j/PHP_XLSXWriter)

## install 

```shell

composer require zhaqq/xlsx

```

##  CLI Example

```php

    $writer = new Builder();
    $fileName = __DIR__ . '/data/xlsx_writer' . date('Ymd-His') . '.xlsx';
    $writer->buildHeader('sheet_name_1', [
         'title' => 'string',
         'content' => 'string',  
         'weight' => 'number',  
    ]);
    $writer->buildHeader('sheet_name_2', [
                 'title' => 'string',
                 'content' => 'string',  
                 'price' => 'price',  
    ]);
    foreach (rows(100) as $row) {
        $writer->writeSheetRow($row[0], $row[1]);
    }
    $writer->writeToFile($fileName);
    
    function rows($n = 100)
    {
        for ($i = 0; $i < $n; $i++) {
            if ($i % 2) {
                yield ['sheet_name_1', [
                    'title' . $i,
                    'content' . $i,
                    $i++,
                ]];
            } else {
                yield ['sheet_name_2', [
                    'title' . $i,
                    'content' . $i,
                    $i++,
                ]];
            }
        }
    }

```
| rows   | time | memory |
| ------ | ---- | ------ |
|  100 | 0.017s | 2MB    |
|  1000 | 0.018s | 2MB    |
|  5000 | 0.151s | 2MB    |
|  50000 | 0.696s | 2MB    |
| 100000 | 1.411s | 2MB    |
| 150000 | 2.067s | 2MB    |
| 200000 | 2.720s | 2MB    |
| 250000 | 3.307s | 2MB    |



## usage

```php

<?php

require __DIR__ . '/vendor/autoload.php';

use Zhaqq\Xlsx\Writer\Builder;

date_default_timezone_set('PRC');

try {
    $writer = new Builder();
    $fileName = __DIR__ . '/data/xlsx_writer' . date('Ymd-His') . '.xlsx';
    $writer->buildHeader('sheet_name_1', [
         'title' => 'string',
         'content' => 'string',  
         'weight' => 'number',  
    ]);
    $writer->buildHeader('sheet_name_2', [
                 'title' => 'string',
                 'content' => 'string',  
                 'price' => 'price',  
    ]);

    foreach (rows() as $row) {
        $writer->writeSheetRow($row[0], $row[1]);
    }

    $writer->writeToFile($fileName);

} catch (\Exception $exception) {
    var_dump($exception->getMessage());
}

function rows($n = 100)
{
    for ($i = 0; $i < $n; $i++) {
        if ($i % 2) {
            yield ['sheet_name_1', [
                'title' . $i,
                'content' . $i,
                $i++,
            ]];
        } else {
            yield ['sheet_name_2', [
                'title' . $i,
                'content' . $i,
                $i++,
            ]];
        }
    }
}

```

## config

*  cell formats

| simple formats | format code |
| ---------- | ---- |
| string   | @ |
| integer  | 0 |
| date     | YYYY-MM-DD |
| datetime | YYYY-MM-DD HH:MM:SS |
| float3    | #,###0.000 |
| price    | #,##0.00 |
| dollar   | [$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00 |
| euro     | #,##0.00 [$€-407];[RED]-#,##0.00 [$€-407] |

More Configuration Information [PHP_XLSXWriter](https://github.com/mk-j/PHP_XLSXWriter)

## License MIT