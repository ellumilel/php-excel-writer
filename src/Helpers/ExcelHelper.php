<?php
namespace Ellumilel\Helpers;

/**
 * Class PHPExcelHelper
 * @package Ellumilel\Helpers
 */
class ExcelHelper
{
    /**
     * @link http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
     * XFE1048577
     */
    const EXCEL_MAX_ROW = 1048576;
    const EXCEL_MAX_RANGE = 2147483647;
    const EXCEL_MAX_COL = 16384;

    /**
     * @param string $dateInput
     *
     * @return mixed
     */
    public static function convertDateTime($dateInput)
    {
        $epoch = 1900;
        $norm = 300;
        $year = $month = $day = $offset = $seconds = 0;
        $dateTime = $dateInput;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $dateTime, $matches)) {
            $year = $matches[1];
            $month = $matches[2];
            $day = $matches[3];
        }

        if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $dateTime, $matches)) {
            $seconds = ($matches[1] * 60 * 60 + $matches[2] * 60 + $matches[3]) / 86400;
        }

        if ("$year-$month-$day" == '1899-12-31' || "$year-$month-$day" == '1900-01-00') {
            return $seconds;
        }
        if ("$year-$month-$day" == '1900-02-29') {
            return 60 + $seconds;
        }
        $range = $year - $epoch;
        // check leapDay
        $leap = (new \DateTime($dateInput))->format('L');
        $mDays = [31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        if (($year < 1900 || $year > 9999) || ($month < 1 || $month > 12) || $day < 1 || $day > $mDays[$month - 1]) {
            return 0;
        }

        $days = $day + ($range * 365) + array_sum(array_slice($mDays, 0, $month - 1));
        $days += intval(($range) / 4) - intval(($range + $offset) / 100);
        $days += intval(($range + $offset + $norm) / 400) - intval($leap);
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }

    /**
     * @param $val
     *
     * @return mixed
     */
    public static function xmlspecialchars($val)
    {
        return str_replace("'", "&#39;", htmlspecialchars($val));
    }

    /**
     * @param int $rowNumber
     * @param int $columnNumber
     *
     * @return string Cell label/coordinates (A1, C3, AA42)
     */
    public static function xlsCell($rowNumber, $columnNumber)
    {
        $n = $columnNumber;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41).$r;
        }

        return $r.($rowNumber + 1);
    }

    /**
     * @link https://msdn.microsoft.com/ru-RU/library/aa365247%28VS.85%29.aspx
     *
     * @param string $filename
     *
     * @return mixed
     */
    public static function checkFilename($filename)
    {
        $invalidCharacter = array_merge(
            array_map('chr', range(0, 31)),
            ['<', '>', '?', '"', ':', '|', '\\', '/', '*', '&']
        );

        return str_replace($invalidCharacter, '', $filename);
    }

    /**
     * @todo  check escaping
     *
     * @param string $cellFormat
     *
     * @return string
     */
    public static function escapeCellFormat($cellFormat)
    {
        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($cellFormat); $i < $ix; $i++) {
            $c = $cellFormat[$i];
            if ($ignoreUntil == '' && $c == '[') {
                $ignoreUntil = ']';
            } else {
                if ($ignoreUntil == '' && $c == '"') {
                    $ignoreUntil = '"';
                } else {
                    if ($ignoreUntil == $c) {
                        $ignoreUntil = '';
                    }
                }
            }
            if ($ignoreUntil == '' &&
                ($c == ' ' || $c == '-' || $c == '(' || $c == ')') &&
                ($i == 0 || $cellFormat[$i - 1] != '_')
            ) {
                $escaped .= "\\".$c;
            } else {
                $escaped .= $c;
            }
        }

        return $escaped;
    }
}
