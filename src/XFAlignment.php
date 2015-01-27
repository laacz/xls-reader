<?php

namespace laacz\XLSParser;

class XFAlignment
{

    ##
    # Values: section 6.115 (p 214) of OOo docs
    public $hor_align = 0;
    ##
    # Values: section 6.115 (p 215) of OOo docs
    public $vert_align = 0;
    ##
    # Values: section 6.115 (p 215) of OOo docs.<br />
    # Note: file versions BIFF7 and earlier use the documented
    # "orientation" attribute; this will be mapped (without loss)
    # into "rotation".
    public $rotation = 0;
    ##
    # 1 = text is wrapped at right margin
    public $text_wrapped = 0;
    ##
    # A number in range(15).
    public $indent_level = 0;
    ##
    # 1 = shrink font size to fit text into cell.
    public $shrink_to_fit = 0;
    ##
    # 0 = according to context; 1 = left-to-right; 2 = right-to-left
    public $text_direction = 0;
}