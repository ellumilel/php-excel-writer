<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$data1 = [
    ['head1', 'head2'],
    ['string1', 'string2'],
    ['string11', 'string22'],
    ['string111', 'string222'],
];
$data2 = [
    ['head1', 'head2', 'head3', 'head4'],
    ['1', '2', '3', '4'],
];
$wExcel = new Ellumilel\ExcelWriter();
$wExcel->setAuthor('Tester');
$wExcel->writeSheet($data1, 'Sheet11');
$wExcel->writeSheet($data2, 'Sheet22');
$wExcel->writeToFile("output_more_tan_one.xlsx");
$time = microtime(true) - $start;
printf("Complete after: %.4F sec.\n", $time);
