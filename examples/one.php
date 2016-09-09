<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$header = [
    'test1' => 'date',
    'test2' => 'string',
    'test3' => 'euro',
    'test4' => 'dollar',
    'test5' => 'float',
    'test6' => 'float_with_sep',
    'test7' => 'string',
];
$wExcel = new Ellumilel\ExcelWriter();
$wExcel->setAuthor('Tester');
$wExcel->writeSheetHeader('Sheet1', $header);
for ($j = 0; $j < 100; $j++) {
    $wExcel->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d'),
        rand(1000, 10000),
        rand(1000, 10000),
        rand(1000, 10000),
        rand(1000, 10000),
        rand(1000, 10000),
        '=HYPERLINK("http://yandex.ru/asd'.rand(1000, 10000).'/sdf='.rand(1000, 10000).'","ссылка")',
    ]);
}
$wExcel->writeToFile("output_one.xlsx");
$time = microtime(true) - $start;
printf("Complete after: %.4F sec.\n", $time);
