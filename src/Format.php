<?php

namespace laacz\XLSParser;

class Format
{
    ##
    # The key into Book.format_map
    public $format_key = 0;
    ##
    # A classification that has been inferred from the format string.
    # Currently, this is used only to distinguish between numbers and dates.
    # <br />Values:
    # <br />FUN = 0 # unknown
    # <br />FDT = 1 # date
    # <br />FNU = 2 # number
    # <br />FGE = 3 # general
    # <br />FTX = 4 # text
    public $type = FUN;
    ##
    # The format string
    public $format_str = '';

    public function __construct($format_key, $ty, $format_str)
    {
        $this->format_key = $format_key;
        $this->type = $ty;
        $this->format_str = $format_str;
    }
}