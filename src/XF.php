<?php

namespace laacz\XLSParser;

class XF
{
    ##
    # 0 = cell XF, 1 = style XF
    public $is_style = 0;
    ##
    # cell XF: Index into Book.xf_list
    # of this XF's style XF<br />
    # style XF: 0xFFF
    public $parent_style_index = 0;
    ##
    #
    public $format_flag = 0;
    ##
    #
    public $font_flag = 0;
    ##
    #
    public $alignment_flag = 0;
    ##
    #
    public $border_flag = 0;
    ##
    #
    public $background_flag = 0;
    ##
    # &nbsp;
    public $protection_flag = 0;
    ##
    # Index into Book.xf_list
    public $xf_index = 0;
    ##
    # Index into Book.font_list
    public $font_index = 0;
    ##
    # Key into Book.format_map
    # <p>
    # Warning: OOo docs on the XF record call this "Index to FORMAT record".
    # It is not an index in the Python sense. It is a key to a map.
    # It is true <i>only</i> for Excel 4.0 and earlier files
    # that the key into format_map from an XF instance
    # is the same as the index into format_list, and <i>only</i>
    # if the index is less than 164.
    # </p>
    public $format_key = 0;
    ##
    # An instance of an XFProtection object.
    public $protection = Null;
    ##
    # An instance of an XFBackground object.
    public $background = Null;
    ##
    # An instance of an XFAlignment object.
    public $alignment = Null;
    ##
    # An instance of an XFBorder object.
    public $border = Null;

}