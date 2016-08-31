<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$header = [
    'test1' => 'datetime',
    'test2' => 'string',
    'test3' => 'euro',
    'test4' => 'dollar',
];
$writer = new Ellumilel\ExcelWriter();
$writer->setAuthor('Tester');
$writer->writeSheetHeader('Sheet1', $header);
for ($i = 0; $i < 500; $i++) {
    $writer->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        rand(0, 1000),
        rand(0, 1000),
        rand(0, 1000),
    ]);
}
$writer->writeToFile("output_one.xlsx");
$time = microtime(true) - $start;
printf("Complete after: %.4F sec.\n", $time);
