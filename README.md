### Big data Excel writer. Relatively low memory usage.
Excel spreadsheet in with (Office 2007+) xlsx format, with just basic features

#### Build:
[![Build Status](https://travis-ci.org/ellumilel/php-excel-writer.svg?branch=master)](http://travis-ci.org/ellumilel/php-excel-writer)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/ellumilel/php-excel-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/ellumilel/php-excel-writer/?branch=master)

#### Use:
- `ZipArchive`, based on PHP's [Zip extension](http://fr.php.net/manual/en/book.zip.php)

#### Supports
* supports PHP 5.4+
* supports simple formulas
* supports currency/date/numeric cell formatting
* takes UTF-8 encoded input
* multiple worksheets

#### Installation
The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

### DevTools
* PHPUnit
* Optional: PHP_CodeSniffer for PSR-X-compatibility checks

Either run

```
php composer.phar require --prefer-dist ellumilel/php-excel-writer
```

or add

```
"ellumilel/php-excel-writer": "*"
```

to the require section of your `composer.json` file.
#### Example
```
$header = [
    'test1' => 'YYYY-MM-DD HH:MM:SS',
    'test2' => 'string',
    'test3' => 'string',
    'test4' => 'string',
    'test5' => 'string',
    'test6' => 'money',
];

$writer = new Ellumilel\ExcelWriter();
$writer->writeSheetHeader('Sheet1', $header);

for ($i = 0; $i < 5000; $i++) {
    $writer->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        rand(0, 1000),
        rand(0, 1000),
        rand(0, 1000),
        rand(0, 1000),
        rand(0, 1000),
    ]);
}

$writer->writeToFile("example.xlsx");
```
