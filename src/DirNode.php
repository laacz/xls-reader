<?php

namespace laacz\XLSParser;

class DirNode
{

    public $tot_size = 0;

    function __construct($DID, $dent)
    {
        $this->DID = $DID;
        # <HBBiii
        list($cbufsize, $this->etype, $this->colour, $this->left_DID, $this->right_DID, $this->root_DID) = array_values(unpack("va/C2b/i3c", substr($dent, 64, 16)));
        # <ii
        list($this->first_SID, $this->tot_size) = array_values(unpack("i2a", substr($dent, 116, 8)));
        if ($cbufsize == 0) {
            $this->name = '';
        } else {
            $this->name = mb_convert_encoding(substr($dent, 0, $cbufsize - 2), 'UTF-8', 'UTF-16LE');
        }
        $this->children = []; # filled in later
        $this->parent = -1; # indicates orphan; fixed up later
        # <IIII
        $this->tsinfo = array_values(unpack('Va/Vb/Vc/Vd', substr($dent, 100, 16)));
    }
}
