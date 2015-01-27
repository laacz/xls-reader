<?php

namespace laacz\XLSParser;


class Note {
    ##
    # Author of note
    public $author = '';
    ##
    # True if the containing column is hidden
    public $col_hidden = 0;
    ##
    # Column index
    public $colx = 0;
    ##
    # List of (offset_in_string, font_index) tuples.
    # Unlike Sheet.{@link #Sheet.rich_text_runlist_map}, the first offset should always be 0.
    public $rich_text_runlist = Null;
    ##
    # True if the containing row is hidden
    public $row_hidden = 0;
    ##
    # Row index
    public $rowx = 0;
    ##
    # True if note is always shown
    public $show = 0;
    ##
    # Text of the note
    public $text = '';

    public $_object_id;

}