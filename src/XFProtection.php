<?php

namespace laacz\XLSParser;

class XFProtection
{

    ##
    # 1 = Cell is prevented from being changed, moved, resized, or deleted
    # (only if the sheet is protected).
    public $cell_locked = 0;
    ##
    # 1 = Hide formula so that it doesn't appear in the formula bar when
    # the cell is selected (only if the sheet is protected).
    public $formula_hidden = 0;

}