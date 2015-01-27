<?php

namespace laacz\XLSParser;

class Colinfo
{
    ##
    # Width of the column in 1/256 of the width of the zero character,
    # using default font (first FONT record in the file).
    public $width = 0;
    ##
    # XF index to be used for formatting empty cells.
    public $xf_index = -1;
    ##
    # 1 = column is hidden
    public $hidden = 0;
    ##
    # Value of a 1-bit flag whose purpose is unknown
    # but is often seen set to 1
    public $bit1_flag = 0;
    ##
    # Outline level of the column, in range(7).
    # (0 = no outline)
    public $outline_level = 0;
    ##
    # 1 = column is collapsed
    public $collapsed = 0;

}