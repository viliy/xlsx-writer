# xlsx - writer

[![Build Status](https://travis-ci.org/viliy/xlsx-writer.svg?branch=master)](https://travis-ci.org/viliy/xlsx-writer)
[![StyleCI](https://github.styleci.io/repos/195766403/shield?branch=master)](https://github.styleci.io/repos/195766403)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/?branch=master)
[![Code Coverage](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/coverage.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/?branch=master)
[![Build Status](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/build.png?b=master)](https://scrutinizer-ci.com/g/viliy/xlsx-writer/build-status/master)
[![Code Intelligence Status](https://scrutinizer-ci.com/g/viliy/xlsx-writer/badges/code-intelligence.svg?b=master)](https://scrutinizer-ci.com/code-intelligence)


## install 

```shell

composer require zhaqq/xlsx

```

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