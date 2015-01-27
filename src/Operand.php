<?php

namespace laacz\XLSParser;


class Operand {
    ##
    # None means that the actual value of the operand is a variable
    # (depends on cell data), not a constant.
    public $value = Null;
    ##
    # oUNK means that the kind of operand is not known unambiguously.
    public $kind = oUNK;
    ##
    # The reconstituted text of the original formula. Function names will be
    # in English irrespective of the original language, which doesn't seem
    # to be recorded anywhere. The separator is ",", not ";" or whatever else
    # might be more appropriate for the end-user's locale; patches welcome.
    public $text = '?';
}