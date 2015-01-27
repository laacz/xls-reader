<?php

namespace laacz\XLSParser;


class Hyperlink {

    ##
    # Index of first row
    public $frowx = Null;
    ##
    # Index of last row
    public $lrowx = Null;
    ##
    # Index of first column
    public $fcolx = Null;
    ##
    # Index of last column
    public $lcolx = Null;
    ##
    # Type of hyperlink. Unicode string, one of 'url', 'unc',
    # 'local file', 'workbook', 'unknown'
    public $type = Null;
    ##
    # The URL or file-path, depending in the type. Unicode string, except
    # in the rare case of a local but non-existent file with non-ASCII
    # characters in the name, in which case only the "8.3" filename is available,
    # as a bytes (3.x) or str (2.x) string, <i>with unknown encoding.</i>
    public $url_or_path = Null;
    ##
    # Description ... this is displayed in the cell,
    # and should be identical to the cell value. Unicode string, or Null. It seems
    # impossible NOT to have a description created by the Excel UI.
    public $desc = Null;
    ##
    # Target frame. Unicode string. Note: I have not seen a case of this.
    # It seems impossible to create one in the Excel UI.
    public $target = Null;
    ##
    # "Textmark": the piece after the "#" in
    # "http://docs.python.org/library#struct_module", or the Sheet1!A1:Z99
    # part when type is "workbook".
    public $textmark = Null;
    ##
    # The text of the "quick tip" displayed when the cursor
    # hovers over the hyperlink.
    public $quicktip = Null;
}