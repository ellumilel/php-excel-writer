<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
$header = [
    'head1' => 'YYYY-MM-DD HH:MM:SS',
    'head2' => 'string',
    'head3' => 'string',
    'head4' => 'string',
    'head5' => 'string',
    'head6' => 'string',
    'head7' => 'string',
    'head8' => 'string',
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
$time = microtime(true) - $start;
echo round(memory_get_usage() / 1048576, 2)." megabytes";
printf("Complete after %.4F sec.\n", $time);
