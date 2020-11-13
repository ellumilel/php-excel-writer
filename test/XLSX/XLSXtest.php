<?php
require_once __DIR__ . '/../../vendor/autoload.php';

use Ellumilel\ExcelWriter;

class XLSXTest extends PHPUnit\Framework\TestCase
{
    public function testConstruct()
    {
        $output_file = __DIR__ . "/unit_test_output_one.xlsx";

        if(file_exists($output_file)){
          unlink($output_file);
        }
        $this->assertFalse(file_exists($output_file));

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
        $writer->writeToFile($output_file);


        $this->assertEquals(true, file_exists($output_file));

        unlink($output_file);
    }
}
