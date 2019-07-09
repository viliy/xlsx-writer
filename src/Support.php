<?php

/**
 * Class Support.
 */

namespace Zhaqq\Xlsx;

/**
 * Class Support.
 */
class Support
{
    /**
     * @param int  $rowNumber
     * @param int  $columnNumber
     * @param bool $absolute
     *
     * @return string
     */
    public static function xlsCell(int $rowNumber, int $columnNumber, bool $absolute = false)
    {
        $n = $columnNumber;
        for ($r = ''; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41).$r;
        }
        if ($absolute) {
            return '$'.$r.'$'.($rowNumber + 1);
        }

        return $r.($rowNumber + 1);
    }

    /**
     * @param $numFormat
     *
     * @return string
     */
    public static function numberFormatStandardized($numFormat)
    {
        if ('money' == $numFormat) {
            $numFormat = 'dollar';
        }
        if ('number' == $numFormat) {
            $numFormat = 'integer';
        }
        switch ($numFormat) {
            case 'string':
                $numFormat = '@';

                break;
            case 'integer':
                $numFormat = '0';

                break;
            case 'date':
                $numFormat = 'YYYY-MM-DD';

                break;
            case 'datetime':
                $numFormat = 'YYYY-MM-DD HH:MM:SS';

                break;
            case 'price':
                $numFormat = '#,##0.00';

                break;
            case 'float3':
                $numFormat = '#,###0.000';

                break;
            case 'dollar':
                $numFormat = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';

                break;
            case 'euro':
                $numFormat = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';

                break;
        }
        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; ++$i) {
            $c = $numFormat[$i];
            if ('' == $ignoreUntil && '[' == $c) {
                $ignoreUntil = ']';
            } elseif ('' == $ignoreUntil && '"' == $c) {
                $ignoreUntil = '"';
            } elseif ($ignoreUntil == $c) {
                $ignoreUntil = '';
            }
            if ('' == $ignoreUntil && (' ' == $c || '-' == $c || '(' == $c || ')' == $c) && (0 == $i || '_' != $numFormat[$i - 1])) {
                $escaped .= '\\'.$c;
            } else {
                $escaped .= $c;
            }
        }

        return $escaped;
    }

    public static function determineNumberFormatType($numFormat)
    {
        $numFormat = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", '', $numFormat);

        switch ($numFormat) {
            case 'GENERAL':
                return 'n_auto';
            case '@':
                return 'n_string';
            case '0':
                return 'n_numeric';
            case 0 != preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat):
            case 0 != preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat):
                return 'n_datetime';
            case 0 != preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat):
            case 0 != preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat):
            case 0 != preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat):
                return 'n_date';
            case 0 != preg_match('/$(?![^"]*+")/', $numFormat):
            case 0 != preg_match('/%(?![^"]*+")/', $numFormat):
            case 0 != preg_match('/0(?![^"]*+")/', $numFormat):
                return 'n_numeric';
            default:
                return 'n_auto';
        }
    }

    /**
     * @param $haystack
     * @param $needle
     *
     * @return false|int|string
     */
    public static function add2listGetIndex(&$haystack, $needle)
    {
        $existingIdx = array_search($needle, $haystack, $strict = true);
        if (false === $existingIdx) {
            $existingIdx = count($haystack);
            $haystack[] = $needle;
        }

        return $existingIdx;
    }

    /**
     * @param $val
     *
     * @return string
     */
    public static function xmlSpecialChars($val)
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodchars = '                              ';

        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badchars, $goodchars); //strtr appears to be faster than str_replace
    }

    /**
     * @param $dateInput
     *
     * @return float|int
     */
    public static function convertDateTime($dateInput)
    {
        $seconds = 0;    // Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;

        $date_time = $dateInput;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }
        unset($junk);

        //using 1900 as epoch, not 1904, ignoring 1904 special case
        // Special cases for Excel.
        if ('1899-12-31' == "$year-$month-$day") {
            return $seconds;
        }    // Excel 1900 epoch
        if ('1900-01-00' == "$year-$month-$day") {
            return $seconds;
        }    // Excel 1900 epoch
        if ('1900-02-29' == "$year-$month-$day") {
            return 60 + $seconds;
        }    // Excel false leapday
        // We calculate the date by calculating the number of days since the epoch
        // and adjust for the number of leap days. We calculate the number of leap
        // days by normalising the year in relation to the epoch. Thus the year 2000
        // becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;
        // Set month days and check for leap year.
        $leap = ((0 == $year % 400) || ((0 == $year % 4) && ($year % 100))) ? 1 : 0;
        $mdays = [31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        // Some boundary checks
        if ($year < $epoch || $year > 9999) {
            return 0;
        }
        if ($month < 1 || $month > 12) {
            return 0;
        }
        if ($day < 1 || $day > $mdays[$month - 1]) {
            return 0;
        }
        // Accumulate the number of days since the epoch.
        $days = $day;    // Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    // Add days for past months
        $days += $range * 365;                      // Add days for past years
        $days += intval(($range) / 4);             // Add leapdays
        $days -= intval(($range + $offset) / 100); // Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400);  // Add 400 year leapdays
        $days -= $leap;                                      // Already counted above
        // Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            ++$days;
        }

        return $days + $seconds;
    }

    public static function sanitizeSheetname($sheetname)
    {
        static $badchars = '\\/?*:[]';
        static $goodchars = '        ';
        $sheetname = strtr($sheetname, $badchars, $goodchars);
        $sheetname = mb_substr($sheetname, 0, 31);
        $sheetname = trim(trim(trim($sheetname), "'")); //trim before and after trimming single quotes

        return !empty($sheetname) ? $sheetname : 'Sheet'.((rand() % 900) + 100);
    }
}
