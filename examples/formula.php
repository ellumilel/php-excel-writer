<?php
require_once __DIR__ . '/../vendor/autoload.php';

$start = microtime(true);
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
$time = microtime(true) - $start;
printf("Complete after: %.4F sec.\n", $time);
