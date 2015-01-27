<?php

namespace laacz\XLSParser;

class XFBorder
{
    ##
    # The colour index for the cell's top line
    public $top_colour_index = 0;
    ##
    # The colour index for the cell's bottom line
    public $bottom_colour_index = 0;
    ##
    # The colour index for the cell's left line
    public $left_colour_index = 0;
    ##
    # The colour index for the cell's right line
    public $right_colour_index = 0;
    ##
    # The colour index for the cell's diagonal lines, if any
    public $diag_colour_index = 0;
    ##
    # The line style for the cell's top line
    public $top_line_style = 0;
    ##
    # The line style for the cell's bottom line
    public $bottom_line_style = 0;
    ##
    # The line style for the cell's left line
    public $left_line_style = 0;
    ##
    # The line style for the cell's right line
    public $right_line_style = 0;
    ##
    # The line style for the cell's diagonal lines, if any
    public $diag_line_style = 0;
    ##
    # 1 = draw a diagonal from top left to bottom right
    public $diag_down = 0;
    ##
    # 1 = draw a diagonal from bottom left to top right
    public $diag_up = 0;

}