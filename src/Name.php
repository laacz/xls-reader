<?php

namespace laacz\XLSParser;

class Name extends Dumpable
{
//    private $repr_these = ['stack'];
    /**
     * @var null|Book
     */
    public $book = Null; # parent

    ##
    # 0 = Visible; 1 = Hidden
    public $hidden = 0;

    ##
    # 0 = Command macro; 1 = Function macro. Relevant only if macro == 1
    public $func = 0;

    ##
    # 0 = Sheet macro; 1 = VisualBasic macro. Relevant only if macro == 1
    public $vbasic = 0;

    ##
    # 0 = Standard name; 1 = Macro name
    public $macro = 0;

    ##
    # 0 = Simple formula; 1 = Complex formula (array formula or user defined)<br />
    # <i>No examples have been sighted.</i>
    public $complex = 0;

    ##
    # 0 = User-defined name; 1 = Built-in name
    # (common examples: Print_Area, Print_Titles; see OOo docs for full list)
    public $builtin = 0;

    ##
    # Function group. Relevant only if macro == 1; see OOo docs for values.
    public $funcgroup = 0;

    ##
    # 0 = Formula definition; 1 = Binary data<br />  <i>No examples have been sighted.</i>
    public $binary = 0;

    ##
    # The index of this object in book.name_obj_list
    public $name_index = 0;

    ##
    # A Unicode string. If builtin, decoded as per OOo docs.
    public $name = "";

    ##
    # An 8-bit string.
    public $raw_formula = '';

    ##
    # -1: The name is global (visible in all calculation sheets).<br />
    # -2: The name belongs to a macro sheet or VBA sheet.<br />
    # -3: The name is invalid.<br />
    # 0 <= scope < book.nsheets: The name is local to the sheet whose index is scope.
    public $scope = -1;

    ##
    # The result of evaluating the formula, if any.
    # If no formula, or evaluation of the formula encountered problems,
    # the result is None. Otherwise the result is a single instance of the
    # Operand class.
    #
    /**
     * @var Operand|null
     */
    public $result = Null;

    # These are added since they've been created dynamically
    public $excel_sheet_index;
    public $extn_sheet_num;
    public $basic_formula_len;
    public $evaluated;

    ##
    # This is a convenience method for the frequent use case where the name
    # refers to a single cell.
    # @return An instance of the Cell class.
    # @throws XLRDError The name is not a constant absolute reference
    # to a single cell.
    public function cell()
    {
        $res = $this->result;
        if ($res) {
            # result should be an instance of the Operand class
            $kind = $res->kind;
            $value = $res->value;
            if ($kind == oREF && strlen($value) == 1) {
                $ref3d = $value[0];
                if (0 <= $ref3d->shtxlo
                    && $ref3d->shtxlo == $ref3d->shtxhi - 1
                    && $ref3d->rowxlo == $ref3d->rowxhi - 1
                    && $ref3d->colxlo == $ref3d->colxhi - 1
                ) {
                    $sh = $this->book->getSheetByIndex($ref3d->shtxlo);
                    return $sh->cell($ref3d->rowxlo, $ref3d->colxlo);
                }
            }
        }
        throw new XLSParserException("Not a constant absolute reference to a single cell");
    }
    ##
    # This is a convenience method for the use case where the name
    # refers to one rectangular area in one worksheet.
    # @param clipped If true (the default), the returned rectangle is clipped
    # to fit in (0, sheet.nrows, 0, sheet.ncols) -- it is guaranteed that
    # 0 <= rowxlo <= rowxhi <= sheet.nrows and that the number of usable rows
    # in the area (which may be zero) is rowxhi - rowxlo; likewise for columns.
    # @return a tuple (sheet_object, rowxlo, rowxhi, colxlo, colxhi).
    # @throws XLRDError The name is not a constant absolute reference
    # to a single area in a single sheet.
    function area2d($clipped = True)
    {

        $res = $this->result;
        if ($res) {
            # result should be an instance of the Operand class
            $kind = $res->kind;
            $value = $res->value;
            if ($kind == oREF && strlen($value) == 1) { # only 1 reference
                $ref3d = $value[0];
                if (0 <= $ref3d->shtxlo && $ref3d->shtxlo == $ref3d->shtxhi - 1) { # only 1 usable sheet
                    $sh = $this->book->getSheetByIndex($ref3d->shtxlo);
                    if (!$clipped) {
                        return [$sh, $ref3d->rowxlo, $ref3d->rowxhi, $ref3d->colxlo, $ref3d->colxhi];
                    }
                    $rowxlo = min($ref3d->rowxlo, $sh->nrows);
                    $rowxhi = max($rowxlo, min($ref3d->rowxhi, $sh->nrows));
                    $colxlo = min($ref3d->colxlo, $sh->ncols);
                    $colxhi = max($colxlo, min($ref3d->colxhi, $sh->ncols));
                    assert(0 <= $rowxlo && $rowxlo <= $rowxhi && $rowxhi <= $sh->nrows);
                    assert(0 <= $colxlo && $colxlo <= $colxhi && $colxhi <= $sh->ncols);
                    return [$sh, $rowxlo, $rowxhi, $colxlo, $colxhi];
                }
            }
        }
        throw new XLSParserException("Not a constant absolute reference to a single area in a single sheet");
    }

}