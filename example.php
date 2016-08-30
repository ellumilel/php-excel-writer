<?php
include_once('src/ExcelWriter.php');
include_once('src/Writer.php');
include_once('src/Sheet.php');

ini_set('display_errors', 0);
ini_set('log_errors', 1);
error_reporting(E_ALL & ~E_NOTICE);

$start = microtime(true);

$header = [
    'test1' => 'date',
    'test2' => 'string',
    'test3' => 'string',
    'test4' => 'string',
];
//$header = ['c1' => 'string', 'c2' => 'string', 'c3' => 'string', 'c4' => 'string'];
$filename = "output_r.xlsx";

$writer = new xlsxWriter\ExcelWriter('Sheet1');
$writer->writeSheetHeader('Sheet1', $header);

for ($i = 0; $i < 5000; $i++) {
    $writer->writeSheetRow('Sheet1', ['2016-05-01', rand() % 10000, rand() % 10000, rand() % 10000]);
}

$writer->writeToFile($filename);

$time = microtime(true) - $start;
printf('Скрипт выполнялся %.4F сек.', $time);
