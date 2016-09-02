<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$header = [
    'head1' => 'YYYY-MM-DD HH:MM:SS',
    'head2' => 'string',
    'head3' => 'string',
    'head4' => 'string',
    'head5' => 'string',
];
$wExcel = new Ellumilel\ExcelWriter();
$wExcel->setAuthor('BigDataTester');
$wExcel->writeSheetHeader('Sheet1', $header);
for ($j = 0; $j < 500000; $j++) {
    $wExcel->writeSheetRow('Sheet1', [
        (new DateTime())->format('Y-m-d H:i:s'),
        'foo',
        'baz',
        'your text hear',
        123123,
    ]);
}
$wExcel->writeToFile("output_big_data.xlsx");
$time = microtime(true) - $start;
printf("Complete after %.4F sec.\n", $time);
