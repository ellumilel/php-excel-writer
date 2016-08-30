<?php
require_once __DIR__ . '/../vendor/autoload.php';

use Ellumilel\ExcelWriter;

class XLSXTest extends PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $header = [
            'test1' => 'integer',
            'test2' => 'integer',
            'test3' => 'integer',
            'test4' => 'integer',
        ];
        $writer = new Ellumilel\ExcelWriter();
        $writer->setAuthor('Tester');
        $writer->writeSheetHeader('Sheet1', $header);
        $writer->writeSheetRow('Sheet1', [1, 2, 3, 4]);
        $writer->writeToFile("unit_test_output_one.xlsx");


        $this->assertEquals(true, file_exists(__DIR__ . "/../unit_test_output_one.xlsx"));
    }
}
