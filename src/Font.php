<?php

namespace laacz\XLSParser;

class Font
{
    ##
    # 1 = Characters are bold. Redundant; see "weight" attribute.
    public $bold = 0;
    ##
    # Values: 0 = ANSI Latin, 1 = System default, 2 = Symbol,
    # 77 = Apple Roman,
    # 128 = ANSI Japanese Shift-JIS,
    # 129 = ANSI Korean (Hangul),
    # 130 = ANSI Korean (Johab),
    # 134 = ANSI Chinese Simplified GBK,
    # 136 = ANSI Chinese Traditional BIG5,
    # 161 = ANSI Greek,
    # 162 = ANSI Turkish,
    # 163 = ANSI Vietnamese,
    # 177 = ANSI Hebrew,
    # 178 = ANSI Arabic,
    # 186 = ANSI Baltic,
    # 204 = ANSI Cyrillic,
    # 222 = ANSI Thai,
    # 238 = ANSI Latin II (Central European),
    # 255 = OEM Latin I
    public $character_set = 0;
    ##
    # An explanation of "colour index" is given in the Formatting
    # section at the start of this document.
    public $colour_index = 0;
    ##
    # 1 = Superscript, 2 = Subscript.
    public $escapement = 0;
    ##
    # 0 = None (unknown or don't care)<br />
    # 1 = Roman (variable width, serifed)<br />
    # 2 = Swiss (variable width, sans-serifed)<br />
    # 3 = Modern (fixed width, serifed or sans-serifed)<br />
    # 4 = Script (cursive)<br />
    # 5 = Decorative (specialised, for example Old English, Fraktur)
    public $family = 0;
    ##
    # The 0-based index used to refer to this Font() instance.
    # Note that index 4 is never used; xlrd supplies a dummy place-holder.
    public $font_index = 0;
    ##
    # Height of the font (in twips). A twip = 1/20 of a point.
    public $height = 0;
    ##
    # 1 = Characters are italic.
    public $italic = 0;
    ##
    # The name of the font. Example: u"Arial"
    public $name = "";
    ##
    # 1 = Characters are struck out.
    public $struck_out = 0;
    ##
    # 0 = None<br />
    # 1 = Single;  0x21 (33) = Single accounting<br />
    # 2 = Double;  0x22 (34) = Double accounting
    public $underline_type = 0;
    ##
    # 1 = Characters are underlined. Redundant; see "underline_type" attribute.
    public $underlined = 0;
    ##
    # Font weight (100-1000). Standard values are 400 for normal text
    # and 700 for bold text.
    public $weight = 400;
    ##
    # 1 = Font is outline style (Macintosh only)
    public $outline = 0;
    ##
    # 1 = Font is shadow style (Macintosh only)
    public $shadow = 0;

    # No methods ...

}