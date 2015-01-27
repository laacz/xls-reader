<?php

namespace laacz\XLSParser;

use DateInterval;
use DateTime;

class Cell
{
    public $ctype;
    public $value;
    public $xf_index;
    /**
     * @var Note[]
     */
    public $notes = [];

    function __construct($ctype, $value, $xf_index = Null, $datemode = 0)
    {
        $this->ctype = $ctype;
        $this->value = $value;

        # Return datetime object
        if ($this->ctype == XL_CELL_DATE) {
            $epoch_1904 = '1904-01-01';
            $epoch_1900 = '1899-12-31';
            $epoch_1900_minus_1 = '1899-12-30';

            $value = (float)$value;

            if ($datemode) {
                $epoch = $epoch_1904;
            } else if ($value < 60) {
                $epoch = $epoch_1900;
            } else {
                # Workaround Excel 1900 leap year bug by adjusting the epoch.
                $epoch = $epoch_1900_minus_1;
            }

            # The integer part of the Excel date stores the number of days since
            # the epoch and the fractional part stores the percentage of the day.
            $days = (int)$value;
            $fraction = $value - $days;

            # Get the the integer and decimal seconds in Excel's millisecond resolution.
            $seconds = (int)(round($fraction * 86400000));
            $seconds = (int)$seconds / 1000;

            if ($value < 1) {
                $h = (int)($seconds / 3600);
                $m = (int)(($seconds / 60) % 60);
                $s = $seconds % 60;
                $value = sprintf("%02d:%02d:%02d", $h, $m, $s);
            } else {
                $value = (new DateTime($epoch))->add(DateInterval::createFromDateString(sprintf("%d days %d seconds", $days, $seconds)))->format("Y-m-d H:i:s");
            }

            $this->value = $value;
        }

        $this->xf_index = $xf_index;
    }

    function __toString()
    {
        return (string)$this->value;
    }

}