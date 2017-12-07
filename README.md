## Big data Excel writer. Relatively low memory usage.
Excel spreadsheet in with (Office 2007+) xlsx format, with just basic features

### Build:
[![Latest Stable Version](https://poser.pugx.org/ellumilel/php-excel-writer/v/stable)](https://packagist.org/packages/ellumilel/php-excel-writer)
[![Latest Unstable Version](https://poser.pugx.org/ellumilel/php-excel-writer/v/unstable)](https://packagist.org/packages/ellumilel/php-excel-writer)
[![Build Status](https://travis-ci.org/ellumilel/php-excel-writer.svg?branch=master)](http://travis-ci.org/ellumilel/php-excel-writer)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/ellumilel/php-excel-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/ellumilel/php-excel-writer/?branch=master)
[![License](https://poser.pugx.org/ellumilel/php-excel-writer/license)](https://packagist.org/packages/ellumilel/php-excel-writer)
[![Total Downloads](https://poser.pugx.org/ellumilel/php-excel-writer/downloads)](https://packagist.org/packages/ellumilel/php-excel-writer)
### Use:
- `ZipArchive`, based on PHP's [Zip extension](http://fr.php.net/manual/en/book.zip.php)

### Supports
* supports PHP 5.4+
* supports simple formulas
* supports currency/date/numeric cell formatting
* takes UTF-8 encoded input
* multiple worksheets

### Dev
* PHPUnit
* Optional: PHP_CodeSniffer for PSR-X-compatibility checks

### Installation
The preferred way to install this extension is through [composer](http://getcomposer.org/download/).
Either run

```
php composer.phar require --prefer-dist ellumilel/php-excel-writer
```

or add

```
"ellumilel/php-excel-writer": ">=0.1.3"
```

to the require section of your `composer.json` file.
### Formats
*   'string' = 'GENERAL'
*   'text' = '@'
*   'integer' = '0'
*   'float_with_sep' = '#,##0.00'
*   'float' = '0.00'
*   'date' = 'YYYY-MM-DD'
*   'datetime' = 'YYYY-MM-DD HH:MM:SS'
*   'dollar' = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00'
*   'money' = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00'
*   'euro' = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]'
*   'rub' = '#,##0.00 [$₽-419];[Red]-#,##0.00 [$₽-419]'
*   'NN' = 'DDD'
*   'NNN' = 'DDDD'
*   'NNNN' = 'DDDD", "'

### Examples
#### Simple:
```
$header = [
    'test1' => 'YYYY-MM-DD HH:MM:SS',
    'test2' => 'string',
    'test3' => 'string',
    'test4' => 'string',
    'test5' => 'string',
    'test6' => 'money',
];

$wExcel = new Ellumilel\ExcelWriter();
$wExcel->writeSheetHeader('Sheet1', $header);
$wExcel->setAuthor('Your name here');
for ($i = 0; $i < 5000; $i++) {
    $wExcel->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        rand(100, 10000),
        rand(100, 10000),
        rand(100, 10000),
        rand(100, 10000),
        rand(100, 10000),
    ]);
}

$wExcel->writeToFile("example.xlsx");
```
#### 3.200.000 cell data example, low memory ~0.84 mb, fast write ~120 sec:
```
$header = [
    'head1' => 'YYYY-MM-DD HH:MM:SS',
    'head2' => 'string',
    'head3' => 'string',
    'head4' => 'string',
    'head5' => 'string',
    'head6' => 'string',
    'head7' => 'string',
    'head8' => 'text',
];
$wExcel = new Ellumilel\ExcelWriter();
$wExcel->setAuthor('BigData Tester');
$wExcel->writeSheetHeader('Sheet1', $header);
for ($ex = 0; $ex < 400000; $ex++) {
    $wExcel->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        'foo',
        'baz',
        'your text hear',
        rand(10000, 100000),
        rand(10000, 100000),
        rand(10000, 100000),
        rand(10000, 100000),
    ]);
}
$wExcel->writeToFile("output_big_data.xlsx");
```
#### Advanced formula/format:
```
$header = [
    'created' => 'date',
    'id' => 'integer',
    'count' => '#,##0',
    'amount' => 'dollar',
    'description' => 'string',
    'money' => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
    'sum' => 'dollar',
    'rub' => 'rub',
];
$data = [
    [
        '2016-01-01',
        123,
        1002,
        '103.00',
        'short string',
        '=D2*0.15',
        '=DOLLAR('.rand(10000, 100000).', 2)',
        rand(10000, 100000),
    ],
    [
        '2016-04-12',
        234,
        2045,
        '2.00',
        'loooooong string',
        '=D3*0.15',
        '=DOLLAR('.rand(10000, 100000).', 2)',
        rand(10000, 100000),
    ],
    [
        '2016-02-05',
        45,
        56,
        '56.00',
        'loooooong loooooong string',
        '=D4*0.15',
        '=DOLLAR('.rand(10000, 100000).', 2)',
        rand(10000, 100000),
    ],
    [
        '2016-06-27',
        534,
        107,
        '678.00',
        'loooooong loooooongloooooong string',
        '=D5*0.15',
        '=DOLLAR('.rand(10000, 100000).', 2)',
        rand(10000, 100000),
    ],
];

$wExcel = new Ellumilel\ExcelWriter();
$wExcel->writeSheet($data, 'Sheet1', $header);
$wExcel->writeToFile('formulas.xlsx');
```
