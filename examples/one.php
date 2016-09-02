<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$header = [
    'test1' => 'datetime',
    'test2' => 'string',
    'test3' => 'euro',
    'test4' => 'dollar',
];
$wExcel = new Ellumilel\ExcelWriter();
$wExcel->setAuthor('Tester');
$wExcel->writeSheetHeader('Sheet1', $header);
for ($i = 0; $i < 500; $i++) {
    $wExcel->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        rand(0, 1000),
        rand(0, 1000),
        rand(0, 1000),
    ]);
}
$wExcel->writeToFile("output_one.xlsx");
$time = microtime(true) - $start;
printf("Complete after: %.4F sec.\n", $time);
