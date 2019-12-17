<?php

namespace laacz\XLSParser;

use ArrayAccess;
use ArrayIterator;
use IteratorAggregate;

class Sheet implements ArrayAccess, IteratorAggregate
{
    /**
     * Name of sheet.
     *
     * @var string
     */
    public $name = '';

    /**
     * A reference to the Book object to which this sheet belongs.
     *
     * @var Book|null
     */
    public $book = Null;

    /**
     * @var Cell[][]
     */
    public $cells = [];

    /**
     * Number of rows in sheet. A row index is in range(thesheet.nrows).
     *
     * @var int
     */
    public $nrows = 0;

    /**
     * Nominal number of columns in sheet. It is 1 + the maximum column index
     * found, ignoring trailing empty cells. See also open_workbook(ragged_rows=?)
     * and Sheet.{@link #Sheet.row_len}(row_index).
     *
     * @var int
     */
    public $ncols = 0;

    /**
     * The map from a column index to a {@link #Colinfo} object. Often there is an entry
     * in COLINFO records for all column indexes in range(257).
     * Note that xlrd ignores the entry for the non-existent
     * 257th column. On the other hand, there may be no entry for unused columns.
     * Populated only if open_workbook(formatting_info=True).
     *
     * @var Colinfo[]
     */
    public $colinfo_map = [];

    /**
     * The map from a row index to a {@link #Rowinfo} object. Note that it is possible
     * to have missing entries -- at least one source of XLS files doesn't
     * bother writing ROW records.
     * Populated only if open_workbook(formatting_info=True).
     * @var Rowinfo[]
     */
    public $rowinfo_map = [];

    /**
     * List of address ranges of cells containing column labels.
     * These are set up in Excel by Insert > Name > Labels > Columns.
     *
     * How to deconstruct the list:
     *
     * for crange in thesheet.col_label_ranges:
     *     rlo, rhi, clo, chi = crange
     *     for rx in xrange(rlo, rhi):
     *         for cx in xrange(clo, chi):
     *             print "Column label at (rowx=%d, colx=%d) is %r" \
     *                 (rx, cx, thesheet.cell_value(rx, cx))
     *
     * @var array
     */
    public $col_label_ranges = [];

    /**
     * List of address ranges of cells containing row labels.
     * For more details, see <i>col_label_ranges</i> above.
     *
     * @var array
     */
    public $row_label_ranges = [];

    /**
     * List of address ranges of cells which have been merged.
     * These are set up in Excel by Format > Cells > Alignment, then ticking
     * the "Merge cells" box.
     *
     * Extracted only if open_workbook(formatting_info=True).
     *
     * How to deconstruct the list:
     *
     * for crange in thesheet.merged_cells:
     *     rlo, rhi, clo, chi = crange
     *     for rowx in xrange(rlo, rhi):
     *         for colx in xrange(clo, chi):
     *             # cell (rlo, clo) (the top left one) will carry the data
     *             # and formatting info; the remainder will be recorded as
     *             # blank cells, but a renderer will apply the formatting info
     *             # for the top left cell (e.g. border, pattern) to all cells in
     *             # the range.
     *
     * @var array
     */
    public $merged_cells = [];

    /**
     * Mapping of (rowx, colx) to list of (offset, font_index) tuples. The offset
     * defines where in the string the font begins to be used.
     * Offsets are expected to be in ascending order.
     * If the first offset is not zero, the meaning is that the cell's XF's font should
     * be used from offset 0.
     *
     * This is a sparse mapping. There is no entry for cells that are not formatted with
     * rich text.
     *
     * How to use:
     *
     * runlist = thesheet.rich_text_runlist_map.get((rowx, colx))
     * if runlist:
     *     for offset, font_index in runlist:
     *         # do work here.
     *         pass
     *
     * Populated only if open_workbook(formatting_info=True).
     *
     * @var array
     */
    public $rich_text_runlist_map = [];

    /**
     * Default column width from DEFCOLWIDTH record, else None.
     * From the OOo docs:
     *
     * > Column width in characters, using the width of the zero character
     * > from default font (first FONT record in the file). Excel adds some
     * > extra space to the default width, depending on the default font and
     * > default font size. The algorithm how to exactly calculate the resulting
     * > column width is not known.
     * >
     * > Example: The default width of 8 set in this record results in a column
     * > width of 8.43 using Arial font with a size of 10 points.
     *
     * For the default hierarchy, refer to the {@link #Colinfo} class.
     *
     * @var null
     */
    public $defcolwidth = Null;

    /**
     * Default column width from STANDARDWIDTH record, else None.
     * From the OOo docs:
     *
     * > Default width of the columns in 1/256 of the width of the zero
     * > character, using default font (first FONT record in the file).
     *
     * For the default hierarchy, refer to the {@link #Colinfo} class.
     *
     * @var null
     */
    public $standardwidth = Null;

    /**
     * Default value to be used for a row if there is
     * no ROW record for that row.
     *
     * From the *optional*> DEFAULTROWHEIGHT record.
     *
     * @var null
     */
    public $default_row_height = Null;

    /**
     * Default value to be used for a row if there is
     * no ROW record for that row.
     *
     * From the <i>optional</i> DEFAULTROWHEIGHT record.
     *
     * @var int|null
     */
    public $default_row_height_mismatch = Null;

    /**
     * Default value to be used for a row if there is
     * no ROW record for that row.
     *
     * From the <i>optional</i> DEFAULTROWHEIGHT record.
     *
     * @var int|null
     */
    public $default_row_hidden = Null;

    /**
     * Default value to be used for a row if there is
     * no ROW record for that row.
     *
     * From the <i>optional</i> DEFAULTROWHEIGHT record.
     *
     * @var int|null
     */
    public $default_additional_space_above = Null;

    /**
     * Default value to be used for a row if there is
     * no ROW record for that row.
     *
     * From the <i>optional</i> DEFAULTROWHEIGHT record.
     *
     * @var int|null
     */
    public $default_additional_space_below = Null;

    /**
     * Visibility of the sheet. 0 = visible, 1 = hidden (can be unhidden
     * by user -- Format/Sheet/Unhide), 2 = "very hidden" (can be unhidden
     * only by VBA macro).
     *
     * @var int
     */
    public $visibility = 0;

    /**
     * A 256-element tuple corresponding to the contents of the GCW record for this sheet.
     *
     * If no such record, treat as all bits zero.
     *
     * Applies to BIFF4-7 only. See docs of the {@link #Colinfo} class for discussion.
     *
     * @var array
     */
    public $gcw = []; # (0, ) * 256;

    /**
     * A list of {@link #Hyperlink} objects corresponding to HLINK records found
     * in the worksheet.
     * @var array
     */
    public $hyperlink_list = [];

    /**
     * <p>A sparse mapping from (rowx, colx) to an item in {@link #Sheet.hyperlink_list}.
     * Cells not covered by a hyperlink are not mapped.
     *
     * It is possible using the Excel UI to set up a hyperlink that
     * covers a larger-than-1x1 rectangle of cells.
     *
     * Hyperlink rectangles may overlap (Excel doesn't check).
     *
     * When a multiply-covered cell is clicked on, the hyperlink that is activated
     * (and the one that is mapped here) is the last in hyperlink_list.
     *
     * @var array
     */
    public $hyperlink_map = [];

    /**
     * A sparse mapping from (rowx, colx) to a {@link #Note} object.
     * Cells not containing a note ("comment") are not mapped.
     *
     * @var array
     */
    public $cell_note_map = [];

    /**
     * Number of columns in left pane (frozen panes; for split panes, see comments below in code)
     *
     * @var int
     */
    public $vert_split_pos = 0;

    /**
     * Number of rows in top pane (frozen panes; for split panes, see comments below in code)
     *
     * @var int
     */
    public $horz_split_pos = 0;

    /**
     * Index of first visible row in bottom frozen/split pane
     *
     * @var int
     */
    public $horz_split_first_visible = 0;

    /**
     * Index of first visible column in right frozen/split pane
     *
     * @var int
     */
    public $vert_split_first_visible = 0;

    /**
     * Frozen panes: ignore it. Split panes: explanation and diagrams in OOo docs.
     *
     * @var int
     */
    public $split_active_pane = 0;

    /**
     * Boolean specifying if a PANE record was present, ignore unless you're xlutils.copy
     *
     * @var int
     */
    public $has_pane_record = 0;

    /**
     * A list of the horizontal page breaks in this sheet.
     *
     * Breaks are tuples in the form (index of row after break, start col index, end col index).
     *
     * Populated only if open_workbook(formatting_info=True).
     *
     * @var array
     */
    public $horizontal_page_breaks = [];

    /**
     * A list of the vertical page breaks in this sheet.
     *
     * Breaks are tuples in the form (index of col after break, start row index, end row index).
     *
     * Populated only if open_workbook(formatting_info=True).
     *
     * @var array
     */
    public $vertical_page_breaks = [];
    public $saved_obj_id = Null;

    /**
     * @param Book $book
     * @param $position
     * @param $name
     * @param $number
     */
    function __construct(Book $book, $position, $name, $number)
    {
        $this->book = $book;
        $this->biff_version = $book->biff_version;
        $this->position = $position;
        $this->name = $name;
        $this->number = $number;
        $this->formatting_info = $book->formatting_info;
        $this->ragged_rows = $book->ragged_rows;
        $this->put_cell = [$this, 'put_cell_unragged'];
        $this->xf_index_to_xl_type_map = $book->xf_index_to_xl_type_map;

        $this->bt = [XL_CELL_EMPTY];
        $this->bf = [-1];
        $this->nrows = 0; # actual, including possibly empty cells
        $this->ncols = 0;
        $this->maxdatarowx = -1; # highest rowx containing a non-empty cell
        $this->maxdatacolx = -1; # highest colx containing a non-empty cell
        $this->dimnrows = 0; # as per DIMENSIONS record
        $this->dimncols = 0;
        $this->cell_values = [];
        $this->cell_types = [];
        $this->cell_xf_indexes = [];
        $this->defcolwidth = Null;
        $this->standardwidth = Null;
        $this->default_row_height = Null;
        $this->default_row_height_mismatch = 0;
        $this->default_row_hidden = 0;
        $this->default_additional_space_above = 0;
        $this->default_additional_space_below = 0;
        $this->colinfo_map = [];
        $this->rowinfo_map = [];
        $this->col_label_ranges = [];
        $this->row_label_ranges = [];
        $this->merged_cells = [];
        $this->rich_text_runlist_map = [];
        $this->horizontal_page_breaks = [];
        $this->vertical_page_breaks = [];
        $this->first_visible_rowx = 0;
        $this->first_visible_colx = 0;
        $this->gridline_colour_index = 0x40;
        $this->gridline_colour_rgb = Null; # pre-BIFF8
        $this->hyperlink_list = [];
        $this->hyperlink_map = [];
        $this->cell_note_map = [];
        # Values calculated by xlrd to predict the mag factors that
        # will actually be used by Excel to display your worksheet.
        # Pass these values to xlwt when writing XLS files.
        # Warning 1: Behaviour of OOo Calc and Gnumeric has been observed to differ from Excel's.
        # Warning 2: A value of zero means almost exactly what it says. Your sheet will be
        # displayed as a very tiny speck on the screen. xlwt will reject attempts to set
        # a mag_factor that is not (10 <= mag_factor <= 400).
        $this->cooked_page_break_preview_mag_factor = 60;
        $this->cooked_normal_view_mag_factor = 100;

        # Values (if any) actually stored on the XLS file
        $this->cached_page_break_preview_mag_factor = Null; # from WINDOW2 record
        $this->cached_normal_view_mag_factor = Null; # from WINDOW2 record
        $this->scl_mag_factor = Null; # from SCL record

        $this->ixfe = Null; # BIFF2 only
        $this->cell_attr_to_xfx = []; # BIFF2.0 only

        $this->utter_max_cols = 256;
        $this->first_full_rowx = -1;

        #### Don't initialise this here, use class attribute initialisation.
        #### $this->gcw = (0, ) * 256 ####


        $this->visibility = $book->sheet_visibility[$number]; # from BOUNDSHEET record
        foreach (Defs::$WINDOW2_options as $attr => $defval) {
            $this->$attr = $defval;
        }

        # Redundant in source

        if ($this->biff_version >= 80) {
            $this->utter_max_rows = 65536;
        } else {
            $this->utter_max_rows = 16384;
        }


    }

    /**
     * {@link #Cell} object in the given row and column.
     *
     * @param $rowx
     * @param $colx
     * @return Cell
     */
    function cell($rowx, $colx)
    {
        if ($this->formatting_info) {
            $xfx = $this->cell_xf_index($rowx, $colx);
        } else {
            $xfx = Null;
        }
        return new Cell($this->cell_types[$rowx][$colx], $this->cell_values[$rowx][$colx], $xfx, $this->book->datemode);
    }

    /**
     * XF index of the cell in the given row and column.
     * This is an index into Book.{@link #Book.xf_list}.
     *
     * @param $rowx
     * @param $colx
     * @return int
     * @throws XLSParserException
     */
    function cell_xf_index($rowx, $colx)
    {
        $this->req_fmt_info();
        $xfx = $this->cell_xf_indexes[$rowx][$colx];
        if ($xfx > -1) {
            return $xfx;
        }
        # Check for a row xf_index
        if (isset($this->rowinfo_map[$rowx]->xf_index)) {
            $xfx = $this->rowinfo_map[$rowx]->xf_index;
            return $xfx;
        }
        return 15;
    }

    /**
     * @throws XLSParserException
     */
    function req_fmt_info()
    {
        if (!$this->formatting_info) {
            throw new XLSParserException('Feature requires open_workbook(..., formatting_info=True)');
        }
    }

    /**
     * @param Book $bk
     * @return int
     * @throws XLSParserException
     */
    function read(Book $bk)
    {
        # Unused in code
        # $r1c1 = 0;
        $oldpos = $bk->position;
        $saved_obj_id = Null;

        $bk->position = $this->position;
        $XL_SHRFMLA_ETC_ETC = [XL_SHRFMLA, XL_ARRAY, XL_TABLEOP, XL_TABLEOP2, XL_ARRAY2, XL_TABLEOP_B2];

        $bv = $this->biff_version;

        $fmt_info = $this->formatting_info;

        $do_sst_rich_text = $fmt_info && $bk->rich_text_runlist_map;
        $rowinfo_sharing_dict = [];
        /** @var $txos MSTxo[] */
        $txos = [];
        $eof_found = 0;

        while (true) {
            list($rc, $data_len, $data) = $bk->readRecordParts();
            switch ($rc) {
                case XL_NUMBER:
                    # <HHHd
                    list($rowx, $colx, $xf_index, $d) = array_values(unpack('v3a/db', substr($data, 0, 14)));
                    $this->put_cell($rowx, $colx, Null, $d, $xf_index);
                    break;
                case XL_LABELSST:
                    # <HHHi
                    list($rowx, $colx, $xf_index, $sstindex) = array_values(unpack('v3a/ib', $data));
                    $this->put_cell($rowx, $colx, XL_CELL_TEXT, $bk->sharedstrings[$sstindex], $xf_index);
                    if ($do_sst_rich_text) {
                        if (isset($bk->rich_text_runlist_map[$sstindex])) {
                            $this->rich_text_runlist_map[$rowx][$colx] = $bk->rich_text_runlist_map[$sstindex];
                        }
                    }
                    break;
                case XL_LABEL:
                    # <HHH
                    list($rowx, $colx, $xf_index) = array_values(unpack('v3a', substr($data, 0, 6)));
                    if ($bv < BIFF_FIRST_UNICODE) {
                        $strg = Helper::unpack_string($data, 6, $bk->encoding ?: $bk->deriveEncoding(), 2);
                    } else {
                        $strg = Helper::unpack_unicode($data, 6, 2);
                    }
                    $this->put_cell($rowx, $colx, XL_CELL_TEXT, $strg, $xf_index);
                    break;
                case XL_STRING:
                    # <HHH
                    list($rowx, $colx, $xf_index) = array_values(unpack('v3a', substr($data, 0, 6)));

                    if ($bv < BIFF_FIRST_UNICODE) {
                        list($strg, $pos) = Helper::unpack_string_update_pos($data, 6, $bk->encoding ?: $bk->deriveEncoding(), 2);
                        $nrt = ord($data[$pos]);
                        $pos++;
                        $runlist = [];
                        for ($tmp = 0; $tmp < $nrt; $tmp++) {
                            $runlist[] = array_values(unpack('C2a', substr($data, $pos, 2)));
                            $pos += 2;
                        }
                        assert($pos == strlen($data));
                    } else {
                        list($strg, $pos) = Helper::unpack_unicode_update_pos($data, 6, 2);
                        list($nrt) = array_values(unpack('v', substr($data, $pos, 2)));
                        $pos += 2;
                        $runlist = [];
                        for ($tmp = 0; $tmp < $nrt; $tmp++) {
                            $runlist[] = array_values(unpack('v2a', substr($data, $pos, 4)));
                            $pos += 4;
                        }
                        assert($pos == strlen($data));
                    }
                    $this->put_cell($rowx, $colx, XL_CELL_TEXT, $strg, $xf_index);
                    $this->rich_text_runlist_map[$rowx][$colx] = $runlist;

                    break;
                case XL_RK:
                    # <HHH
                    list($rowx, $colx, $xf_index) = array_values(unpack('v3a', substr($data, 0, 6)));
                    $d = Helper::unpack_RK(substr($data, 6, 4));
                    $this->put_cell($rowx, $colx, Null, $d, $xf_index);
                    break;
                case XL_MULRK:
                    # <HH
                    list($mulrk_row, $mulrk_first) = array_values(unpack('v2', substr($data, 0, 4)));
                    # <H
                    list($mulrk_last) = array_values(unpack('v', substr($data, -2)));

                    $pos = 4;
                    for ($colx = $mulrk_first; $colx <= $mulrk_last; $colx++) {
                        # <H
                        list($xf_index) = array_values(unpack('v', substr($data, $pos, 2)));
                        $d = Helper::unpack_RK(substr($data, $pos + 2, 4));
                        $pos += 6;
                        $this->put_cell($mulrk_row, $colx, Null, $d, $xf_index);
                    }
                    break;
                case XL_ROW:
                    # Version 0.6.0a3: ROW records are just not worth using (for memory allocation).
                    # Version 0.6.1: now used for formatting info.
                    if (!$fmt_info) break;
                    # <H4xH4xi
                    list($rowx, , , , , $bits1, , , , , $bits2) = array_values(unpack('va/c4b/vc/c4d/i', substr($data, 0, 16)));
                    if (!(0 <= $rowx && $rowx < $this->utter_max_rows)) {
                        break;
                    }
                    $r = isset($rowinfo_sharing_dict[$bits1][$bits2]) ? $rowinfo_sharing_dict[$bits1][$bits2] : Null;
                    if ($r == Null) {
                        $r = new Rowinfo();
                        # Using upkbits() is far too slow on a file
                        # with 30 sheets each with 10K rows :-(
                        #    upkbits(r, bits1, (
                        #        ( 0, 0x7FFF, 'height'),
                        #        (15, 0x8000, 'has_default_height'),
                        #        ))
                        #    upkbits(r, bits2, (
                        #        ( 0, 0x00000007, 'outline_level'),
                        #        ( 4, 0x00000010, 'outline_group_starts_ends'),
                        #        ( 5, 0x00000020, 'hidden'),
                        #        ( 6, 0x00000040, 'height_mismatch'),
                        #        ( 7, 0x00000080, 'has_default_xf_index'),
                        #        (16, 0x0FFF0000, 'xf_index'),
                        #        (28, 0x10000000, 'additional_space_above'),
                        #        (29, 0x20000000, 'additional_space_below'),
                        #        ))
                        # So:
                        $r->height = $bits1 & 0x7fff;
                        $r->has_default_height = ($bits1 >> 15) & 1;
                        $r->outline_level = $bits2 & 7;
                        $r->outline_group_starts_ends = ($bits2 >> 4) & 1;
                        $r->hidden = ($bits2 >> 5) & 1;
                        $r->height_mismatch = ($bits2 >> 6) & 1;
                        $r->has_default_xf_index = ($bits2 >> 7) & 1;
                        $r->xf_index = ($bits2 >> 16) & 0xfff;
                        $r->additional_space_above = ($bits2 >> 28) & 1;
                        $r->additional_space_below = ($bits2 >> 29) & 1;
                        if (!$r->has_default_xf_index) {
                            $r->xf_index = -1;
                        }
                        $rowinfo_sharing_dict[$bits1][$bits2] = $r;
                    }
                    $this->rowinfo_map[$rowx] = $r;

                    break;
                case 0x0006: # XL_FORMULA_OPCODES:
                case 0x0406: # XL_FORMULA_OPCODES:
                case 0x0206: # XL_FORMULA_OPCODES:
                    if ($bv >= 50) {
                        # <HHH8sH
                        list($rowx, $colx, $xf_index, $flags) = array_values(unpack('va/vb/vc/x8/ve', substr($data, 0, 16)));
                        $result_str = substr($data, 6, 8);
                        $lenlen = 2;
                        $tkarr_offset = 20;
                    } elseif ($bv >= 30) {
                        # <HHH8sH
                        list($rowx, $colx, $xf_index, $flags) = array_values(unpack('va/vb/vc/x8/ve', substr($data, 0, 16)));
                        $result_str = substr($data, 6, 8);
                        $lenlen = 2;
                        $tkarr_offset = 16;
                    } else { # BIFF2
                        # <HH3s8sB
                        list($rowx, $colx, $flags) = array_values(unpack('va/vb/x11/cc', substr($data, 0, 16)));
                        $cell_attr = substr($data, 4, 3);
                        $result_str = substr($data, 7, 8);
                        $xf_index = $this->fixed_BIFF2_xfindex($cell_attr, $rowx, $colx);
                        $lenlen = 1;
                        $tkarr_offset = 16;
                    }
                    if (substr($result_str, 6, 2) == "\xff\xff") {
                        $first_byte = ord($result_str[0]);
                        switch ($first_byte) {
                            case 0:
                                # need to read next record (STRING)
                                $gotstring = 0;
                                # actually there's an optional SHRFMLA or ARRAY etc record to skip over
                                list($rc2, $data2_len, $data2) = $bk->readRecordParts();
                                switch ($rc2) {
                                    case XL_STRING:
                                    case XL_STRING_B2:
                                        $gotstring = 1;
                                        break;
                                    case XL_ARRAY:
                                        # <HHBBBxxxxxH
//                                        list($row1x, $rownx, $col1x, $colnx, $array_flags, $tokslen) = array_values(unpack("v2a/c3b/x5/vc", substr($data2, 0, 14)));
                                        break;
                                    case XL_SHRFMLA:
                                        # <HHBBxBH
//                                        list($row1x, $rownx, $col1x, $colnx, $nfmlas, $tokslen) = array_values(unpack("v2a/C2b/x/Cc/vd", substr($data, 0, 10)));
                                        break;
                                    default:
                                        if (!in_array($rc2, $XL_SHRFMLA_ETC_ETC)) {
                                            throw new XLSParserException(sprintf("Expected SHRFMLA, ARRAY, TABLEOP* or STRING record; found 0x%04x", $rc2));
                                        }
                                }
                                if (!$gotstring) {
                                    list($rc2, $tmp, $data2) = $bk->readRecordParts();
                                    if ($rc2 !== XL_STRING && $rc2 !== XL_STRING_B2) {
                                        throw new XLSParserException(sprintf("Expected STRING record; found 0x%04x", $rc2));
                                    }
                                }
                                $strg = $this->string_record_contents($data2);
                                $this->put_cell($rowx, $colx, XL_CELL_TEXT, $strg, $xf_index);
                                break;
                            case 1:
                                # boolean formula result
                                $value = ord($result_str[2]);
                                $this->put_cell($rowx, $colx, XL_CELL_BOOLEAN, $value, $xf_index);
                                break;
                            case 2:
                                # Error in cell
                                $value = ord($result_str[2]);
                                $this->put_cell($rowx, $colx, XL_CELL_ERROR, $value, $xf_index);
                                break;
                            case 3:
                                # empty ... i.e. empty (zero-length) string, NOT an empty cell.
                                $this->put_cell($rowx, $colx, XL_CELL_TEXT, "", $xf_index);
                                break;
                            default:
                                throw new XLSParserException(sprintf("unexpected special case (0x%02x) in FORMULA", $first_byte));
                        }
                    } else {
                        # <d
                        list($d) = array_values(unpack("d", $result_str));
                        $this->put_cell($rowx, $colx, Null, $d, $xf_index);
                    }
                    break;
                case XL_BOOLERR:
                    # <HHHBB
                    list($rowx, $colx, $xf_index, $value, $is_err) = array_values(unpack('v3a/c2', substr($data, 0, 8)));
                    # Note OOo Calc 2.0 writes 9-byte BOOLERR records.
                    # OOo docs say 8. Excel writes 8.
                    $cellty = $is_err ? XL_CELL_ERROR : XL_CELL_BOOLEAN;
                    $this->put_cell($rowx, $colx, $cellty, $value, $xf_index);
                    break;
                case XL_COLINFO:
                    if (!$fmt_info) break;
                    $c = new Colinfo();
                    # <HHHHH
                    list($first_colx, $last_colx, $c->width, $c->xf_index, $flags) = array_values(unpack('v5a', substr($data, 0, 10)));
                    #### Colinfo.width is denominated in 256ths of a character,
                    #### *not* in characters.
                    if (!(0 <= $first_colx && $first_colx <= $last_colx && $last_colx <= 256)) {
                        # Note: 256 instead of 255 is a common mistake.
                        # We silently ignore the non-existing 257th column in that case.
                        break;
                    }

                    Helper::upkbits($c, $flags, [
                        [0, 0x0001, 'hidden'],
                        [1, 0x0002, 'bit1_flag'],
                        # *ALL* colinfos created by Excel in "default" cases are 0x0002!!
                        # Maybe it's "locked" by analogy with XFProtection data.
                        [8, 0x0700, 'outline_level'],
                        [12, 0x1000, 'collapsed'],
                    ]);

                    for ($colx = $first_colx; $colx <= $last_colx; $colx++) {
                        if ($colx > 255) break; # Excel does 0 to 256 inclusive
                        $this->colinfo_map[$colx] = $c;
                    }
                    break;
                case XL_DEFCOLWIDTH:
                    # <HHHHH
                    list($this->defcolwidth) = array_values(unpack('v', substr($data, 0, 2)));
                    break;
                case XL_STANDARDWIDTH:
                    if ($data_len == 2) {
                        list($this->standardwidth) = array_values(unpack('v', substr($data, 0, 2)));
                    }
                    break;
                case XL_GCW:
                    if (!$fmt_info) break;
                    assert($data_len == 34);
                    assert(substr($data, 0, 2) == "\x20\x00");
                    # <8i
                    $iguff = array_values(unpack("i8", substr($data, 2, 32)));
                    $gcw = [];
//                    print_r($iguff);
                    foreacH ($iguff as $bits) {
                        for ($j = 0; $j < 32; $j++) {
                            $gow[] = $bits & 1;
                            $bits >>= 1;
                        }
                    }
                    $this->gcw = $gcw;
                    break;
                case XL_BLANK:
                    if (!$fmt_info) break;
                    # <HHH
                    list($rowx, $colx, $xf_index) = array_values(unpack('v3a', substr($data, 0, 6)));
                    $this->put_cell($rowx, $colx, XL_CELL_BLANK, '', $xf_index);
                    break;
                case XL_MULBLANK:
                    if (!$fmt_info) break;
                    $nitems = $data_len >> 1;
                    # "<%dH" % nitems
                    $result = array_values(unpack('v' . $nitems . 'a', $data));
                    $rowx = $result[0];
                    $mul_first = $result[1];
                    $mul_last = $result[count($result) - 1];
                    assert($nitems == $mul_last + 4 - $mul_first);
                    $pos = 2;
                    for ($colx = $mul_first; $colx <= $mul_last; $colx++) {
                        $this->put_cell($rowx, $colx, XL_CELL_BLANK, '', $result[$pos]);
                        $pos++;
                    }
                    break;
                case XL_DIMENSION:
                case XL_DIMENSION2:
                    # Four zero bytes after some other record. See xlrd github issue 64.
                    if ($data_len == 0) break;
                    # if data_len == 10:
                    # Was crashing on BIFF 4.0 file w/o the two trailing unused bytes.
                    # Reported by Ralph Heimburger.
                    if ($bv < 80) {
                        # <HxxH
                        list($this->dimnrows, , , $this->dimncols) = array_values(unpack('va/c2b/vc', substr($data, 2, 6)));
                    } else {
                        # <ixxH
                        list($this->dimnrows, , , $this->dimncols) = array_values(unpack('ia/c2b/vc', substr($data, 4, 8)));
                    }
                    $this->nrows = $this->ncols = 0;
                    if (in_array($bv, [21, 30, 40]) && $this->book->xf_list && !$this->book->xf_epilogue_done) {
                        $this->book->XFEpilogue();
                    }
                    break;
                case XL_HLINK:
                    $this->handle_hlink($data);
                    break;
                case XL_QUICKTIP:
                    $this->handle_quicktip($data);
                    break;
                case XL_EOF:
                    $eof_found = true;
                    break;
                case XL_OBJ:
                    # handle SHEET-level objects; note there's a separate Book.handle_obj
                    $saved_obj = $this->handle_obj($data);
//                    $saved_obj_id = Null;
                    if ($saved_obj) {
                        $saved_obj_id = $saved_obj->id;
                    }
                    break;
                case XL_MSO_DRAWING:
                    $this->handle_msodrawingetc($rc, $data_len, $data);
                    break;
                case XL_TXO:
                    $txo = $this->handle_txo($data);
                    if ($txo && $saved_obj_id) {
                        $txos[$saved_obj_id] = $txo;
                        $saved_obj_id = Null;
                    }
                    break;
                case XL_NOTE:
                    $this->handle_note($data, $txos);
                    break;
                case XL_FEAT11:
                    $this->handle_feat11($data);
                    break;
                case 0x0809: # bofcodes
                case 0x0409: # bofcodes
                case 0x0209: # bofcodes
                case 0x0009: # bofcodes
                    # <HH
                    list($version, $boftype) = array_values(unpack('v2', substr($data, 0, 4)));
                    if ($boftype != 0x20) {
                        Helper::log("*** Unexpected embedded BOF (0x%04x) at offset %d: version=0x%04x type=0x%04x", $rc, $bk->position - $data_len - 4, $version, $boftype);
                    }
                    $done = false;
                    while (!$done) {
                        list($code, $data_len, $data) = $this->book->readRecordParts();
                        if ($code == XL_EOF) {
                            $done = true;
                        }
                    }
                    break;
                case XL_COUNTRY:
                    $bk->handleCountry($data);
                    break;
                case XL_LABELRANGES:
                    // @TODO
                    die(Helper::explain_const($rc, 'XL_'));
                    break;
                case XL_ARRAY:
                    // @TODO
                    # <HHBBBxxxxxH
                    list($row1x, $rownx, $col1x, $colnx, $array_flags, $tokslen) = array_values(unpack("v2a/C3b/x5/vc", substr($data, 0, 14)));
                    break;
                case XL_SHRFMLA:
                    # <HHBBxBH
                    list($row1x, $rownx, $col1x, $colnx, $nfmlas, $tokslen) = array_values(unpack("v2a/C2b/x/Cc/vD", substr($data, 0, 10)));
                    // @TODO ? Or not?
                    break;
                case XL_CONDFMT:
                    if (!$fmt_info) break;
                    assert($bv >= 80);
                    # <6H
                    list($num_CFs, $needs_recalc, $browx1, $browx2, $bcolx1, $bcolx2) = array_values(unpack('v6', substr($data, 0, 12)));
                    $olist = []; # updated by the function
                    $pos = Helper::unpack_cell_range_address_list_update_pos($olist, $data, 12, $bv, 8);
                    break;
                case XL_CF:
                    if (!$fmt_info) break;
                    # <BBHHi
                    list($cf_type, $cmp_op, $sz1, $sz2, $flags) = array_values(unpack("C2a/v2b/ic", substr($data, 0, 10)));
                    $font_block = ($flags >> 26) & 1;
                    $bord_block = ($flags >> 28) & 1;
                    $patt_block = ($flags >> 29) & 1;
                    $pos = 12;
                    if ($font_block) {
                        # <64x i i H H B 3x i 4x i i i 18x
                        list($font_height, $font_options, $weight, $escapement, $underline, $font_colour_index, $two_bits, $font_esc, $font_underl) = array_values(unpack("x64/i2a/v2b/Cc/x3/id/x4/i3e/x18", substr($data, $pos, 118)));
                        $font_style = ($two_bits > 1) & 1;
                        $posture = ($font_options > 1) & 1;
                        $font_canc = ($two_bits > 7) & 1;
                        $cancellation = ($font_options > 7) & 1;

                        $pos += 118;
                    }

                    if ($bord_block) {
                        $pos += 8;
                    }

                    if ($patt_block) {
                        $pos += 8;
                    }

                    $fmla1 = substr($data, $pos, $sz2);
                    $pos += $sz2;
                    assert($pos = $data_len);

//                    die(Helper::explain_const($rc, 'XL_'));
                    break;
                case XL_DEFAULTROWHEIGHT:
                    if ($data_len == 4) {
                        # <HH
                        list($bits, $this->default_row_height) = array_values(unpack('v2a', substr($data, 0, 4)));
                    } elseif ($data_len == 2) {
                        # <H
                        list($this->default_row_height) = array_values(unpack('v', substr($data, 0, 2)));
                        $bits = 0;
                    } else {
                        $bits = 0;
                    }
                    $this->default_row_height_mismatch = $bits & 1;
                    $this->default_row_hidden = ($bits >> 1) & 1;
                    $this->default_additional_space_above = ($bits >> 2) & 1;
                    $this->default_additional_space_below = ($bits >> 3) & 1;

                    break;
                case XL_MERGEDCELLS:
                    if (!$fmt_info) break;
                    $pos = Helper::unpack_cell_range_address_list_update_pos($this->merged_cells, $data, 0, $bv, 8);
                    assert($pos == $data_len);
                    break;
                case XL_WINDOW2:
                    if ($bv >= 80 && $data_len >= 14) {
                        # <HHHHxxHH
                        list($options, $this->first_visible_rowx, $this->first_visible_colx, $this->gridline_colour_index, $this->cached_page_break_preview_mag_factor, $this->cached_normal_view_mag_factor) = array_values(unpack('v7a', substr($data, 0, 14)));
                    } else {
                        assert($bv >= 30); # BIFF3-7
                        # <HHH
                        list($options, $this->first_visible_rowx, $this->first_visible_colx) = array_values(unpack('v3a', substr($data, 0, 6)));
                        # <BBB
                        $this->gridline_colour_rgb = array_values(unpack('C3a', substr($data, 6, 3)));
                        $this->gridline_colour_index = Helper::nearest_colour_index($this->book->colour_map, $this->gridline_colour_rgb);
                        $this->cached_page_break_preview_mag_factor = 0; # default (60%)
                        $this->cached_normal_view_mag_factor = 0; # default (100%)
                    }
                    # options -- Bit, Mask, Contents:
                    # 0 0001H 0 = Show formula results 1 = Show formulas
                    # 1 0002H 0 = Do not show grid lines 1 = Show grid lines
                    # 2 0004H 0 = Do not show sheet headers 1 = Show sheet headers
                    # 3 0008H 0 = Panes are not frozen 1 = Panes are frozen (freeze)
                    # 4 0010H 0 = Show zero values as empty cells 1 = Show zero values
                    # 5 0020H 0 = Manual grid line colour 1 = Automatic grid line colour
                    # 6 0040H 0 = Columns from left to right 1 = Columns from right to left
                    # 7 0080H 0 = Do not show outline symbols 1 = Show outline symbols
                    # 8 0100H 0 = Keep splits if pane freeze is removed 1 = Remove splits if pane freeze is removed
                    # 9 0200H 0 = Sheet not selected 1 = Sheet selected (BIFF5-BIFF8)
                    # 10 0400H 0 = Sheet not visible 1 = Sheet visible (BIFF5-BIFF8)
                    # 11 0800H 0 = Show in normal view 1 = Show in page break preview (BIFF8)
                    # The freeze flag specifies, if a following PANE record (6.71) describes unfrozen or frozen panes.
                    foreach (Defs::$WINDOW2_options as $attr => $defval) {
                        $this->$attr = $options & 1;
                        $options >>= 1;
                    }
                    break;
                case XL_SCL:
                    list($num, $den) = array_values(unpack('v2', $data));
                    $result = 0;
                    if ($den) {
                        $result = (int)($num * 100 / $den);
                    }
                    if (!(10 <= $result && $result <= 40)) {
                        $result = 100;
                    }
                    $this->scl_mag_factor = $result;
                    break;
                case XL_PANE:
                    list($this->vert_split_pos, $this->horz_split_pos, $this->horz_split_first_visible, $this->vert_split_first_visible, $this->split_active_pane) = array_values(unpack('v4a/Cb', substr($data, 0, 9)));
                    $this->has_pane_record = 1;
                    break;
                case XL_HORIZONTALPAGEBREAKS:
                    if (!$fmt_info) break;
                    # <H
                    list($num_breaks) = array_values(unpack('v', substr($data, 0, 2)));
                    assert($num_breaks * (2 + 4 * ($bv >= 80)) + 2 == $data_len);
                    $pos = 2;
                    if ($bv < 80) {
                        while ($pos < $data_len) {
                            # <H
                            list($tmp) = array_values(unpack('v', substr($data, $pos, 2)));
                            $this->horizontal_page_breaks[] = [$tmp, 0, 255];
                            $pos += 2;
                        }
                    } else {
                        while ($pos < $data_len) {
                            # <HHH
                            $this->horizontal_page_breaks[] = array_values(unpack('v3', substr($data, $pos, 6)));
                            $pos += 6;
                        }
                    }
                    break;
                case XL_VERTICALPAGEBREAKS:
                    // @TODO
                    die(Helper::explain_const($rc, 'XL_'));
                    break;
                default:
                    if ($bv <= 45) {
                        #### all of the following are for BIFF <= 4W

                    } else {
//                        $const_name = Helper::explain_const($rc, 'XL_');
                    }
                    break;

            }
            if ($eof_found) {
                break;
            }

        }
        if (!$eof_found) {
            throw new XLSParserException("Sheet {$this->number} ($this->name) missing EOF record");
        }

        $this->tidy_dimensions();
        $this->update_cooked_mag_factors();
        $bk->position = $oldpos;
        return 1;

    }

    public function handle_feat11($data)
    {
        return;
    }

    public function handle_quicktip($data)
    {
        # <5H
        list($rcx, $frowx, $lrowx, $fcolx, $lcolx) = array_values(unpack('v5', substr($data, 0, 10)));
        assert($rcx = XL_QUICKTIP);
        assert($this->hyperlink_list);
        $h = $this->hyperlink_list[count($this->hyperlink_list) - 1];
        assert ([$frowx, $lrowx, $fcolx, $lcolx] == [$h->frowx, $h->lrowx, $h->fcolx, $h->lcolx]);
        assert(substr($data, -2) == "\x00\x00");
        $this->hyperlink_list[count($this->hyperlink_list) - 1]->quicktip = mb_convert_encoding(substr($data, 10, strlen($data) - 2), 'utf-8', 'utf-16le');
    }

    public function handle_hlink($data)
    {
        $record_size = strlen($data);
        $h = new Hyperlink();
        # <HHHH16s4si
        list($h->frowx, $h->lrowx, $h->fcolx, $h->lcolx, $options) = array_values(unpack('v4a/x20/ib', substr($data, 0, 32)));
        $guid0 = substr($data, 8, 16);
        $dummy = substr($data, 24, 4);

        assert($guid0 == "\xD0\xC9\xEA\x79\xF9\xBA\xCE\x11\x8C\x82\x00\xAA\x00\x4B\xA9\x0B");
        assert($dummy == "\x02\x00\x00\x00");

        $offset = 32;
        if ($options & 0x14) {
            list($h->desc, $offset) = Helper::get_nul_terminated_unicode($data, $offset);
        }

        if ($options& 0x80) {
            list($h->target, $offset) = Helper::get_nul_terminated_unicode($data, $offset);
        }

        if ($options & 1 && !($options & 0x100)) { # HasMoniker and not MonikerSavedAsString
            # an OLEMoniker structure
            $clsid = substr($data, $offset, 16);
            $offset += 16;
            if ($clsid == "\xE0\xC9\xEA\x79\xF9\xBA\xCE\x11\x8C\x82\x00\xAA\x00\x4B\xA9\x0B") {
                # URL Moniker
                $h->type = 'url';
                list($nbytes) = array_values(unpack('l', substr($data, $offset, $offset + 4)));
                $offset += 4;
                $h->url_or_path = mb_convert_encoding(substr($data, $offset, $nbytes - 1), 'utf-8', 'utf-16le');
                $offset += $nbytes;

            } else if ($clsid == "\x03\x03\x00\x00\x00\x00\x00\x00\xC0\x00\x00\x00\x00\x00\x00\x46") {
                # file Moniker
                $h->type = 'local file';
                # <Hi
                list($uplevels, $nbytes) = array_values(unpack('va/ib', substr($data, $offset, $offset + 6)));
                $offset += 6;
                $shortpath = mb_convert_encoding(str_repeat("..\\", $uplevels) . substr($data, $offset, $nbytes - 1), 'utf-8');
                $offset += $nbytes;
                $offset += 24; # OOo: "unknown byte sequence"
                # above is version 0xDEAD + 20 reserved zero bytes
                # <i
                list($sz) = array_values(unpack('i', substr($data, $offset, $offset + 4)));
                $offset += 4;
                if ($sz) {
                    # <i
                    list($xl) = array_values(unpack('i', substr($data, $offset, $offset + 4)));
                    $offset += 4;
                    $offset += 2; # "unknown byte sequence" MS: 0x0003
                    $extended_path = substr($data, $offset, $offset + $xl); # not zero-terminated
                    $offset += $xl;
                    $h->url_or_path = $extended_path;
                } else {
                    $h->url_or_path = $shortpath;
                    #### MS KLUDGE WARNING ####
                    # The "shortpath" is bytes encoded in the **UNKNOWN** creator's "ANSI" encoding.
                }
            }
        } else if ($options & 0x163 == 0x103) { # UNC
            $h->type = 'UNC';
            list($h->url_or_path, $offset) = Helper::get_nul_terminated_unicode($data, $offset);
        } else if ($options & 0x16B == 8) {
            $h->type = 'workbook';
        } else {
            $h->type = 'unknown';
        }

        if ($options & 0x8) { # has textmark
            list($h->textmark, $offset) = Helper::get_nul_terminated_unicode($data, $offset);
        }
        if ($record_size - $offset < 0) {
            throw new XLSParserException("Bug or corrupt file, send copy or input file for debugging");
        }

        $this->hyperlink_list[] = $h;
        for ($rowx = $h->lrowx; $rowx <= $h->fcolx; $rowx++) {
            for ($colx = $h->lcolx; $colx <= $h->fcolx; $colx++) {
                $this->hyperlink_map[$rowx][$colx] = $h;
            }
        }

    }

    public function handle_obj($data)
    {
        if ($this->biff_version < 80) {
            return Null;
        }
        $o = new MSObj();
        $data_len = strlen($data);
        $pos = 0;

        while ($pos < $data_len) {
            # <HH
            list($ft, $cb) = array_values(unpack("v2", substr($data, $pos, 4)));
            if ($pos == 0 && !($ft == 0x15 && $cb == 18)) {
                # ignoring antique or corupt OBJECT record!
                return Null;
            }

            if ($ft == 0x15) { # ftCmo ... s/n first
                assert($pos == 0);
                # <HHH
                list($o->type, $o->id, $option_flags) = array_values(unpack("v3", substr($data, $pos + 4, 6)));
                Helper::upkbits($o, $option_flags, [
                    [ 0, 0x0001, 'locked'],
                    [ 4, 0x0010, 'printable'],
                    [ 8, 0x0100, 'autofilter'], # not documented in Excel 97 dev kit
                    [ 9, 0x0200, 'scrollbar_flag'], # not documented in Excel 97 dev kit
                    [13, 0x2000, 'autofill'],
                    [14, 0x4000, 'autoline'],
                ]);
            } elseif ($ft == 0x00) {
                if (substr($data, $pos, $data_len - $pos) == str_repeat("\x00", $data_len - $pos)) {
                    # ignore "optional reserved" data at end of record
                    break;
                }
                throw new XLSParserException("Unexpected data at end of OBJET record");
            } elseif ($ft == 0x0c) { # Scrollbar
                $values = array_values(unpack("v5", substr($data, $pos + 8, 10)));
            } elseif ($ft == 0x0D) {
                # not documented in Excel 97 dev kit
            } elseif ($ft == 0x13) { # list box data
                if ($o->autofilter) { # non standard exit. NOT documented
                    break;
                }
            }


            $pos += $cb + 4;

        }
        return $o;
    }

    public function handle_msodrawingetc($recid, $data_len, $data)
    {
        // TODO: Have no clue what to do with them
        return false;
    }

    public function handle_txo($data)
    {
        if ($this->biff_version < 80) return;
        $o = new MSTxo();
        $data_len = strlen($data);
        # <HH6sHHH
        list($option_flags, $o->rot, $cchText, $cbRuns, $o->ifntEmpty) = array_values(unpack('v2a/x6/v3b', substr($data, 0, 16)));
        $controlInfo = substr($data, 4, 6);
        $o->fmla = substr($data, 16);
        Helper::upkbits($o, $option_flags, [
            [ 3, 0x000E, 'horz_align'],
            [ 6, 0x0070, 'vert_align'],
            [ 9, 0x0200, 'lock_text'],
            [14, 0x4000, 'just_last'],
            [15, 0x8000, 'secret_edit'],
        ]);
        $totchars = 0;
        $o->text = '';
        while ($totchars < $cchText) {
            list($rc2, $data2_len, $data2) = $this->book->readRecordParts();
            assert($rc2 == XL_CONTINUE);
            $nb = ord($data2[0]);
            $nchars = $data2_len - 1;
            if ($nb) {
                assert($nchars % 2 == 0);
                $nchars = (int)$nchars / 2;
            }
            list($utext, $endpos) = Helper::unpack_unicode_update_pos($data2, 0, 2, $nchars);
            assert($endpos == $data2_len);
            $o->text .= $utext;
            $totchars += $nchars;
        }
        $o->rich_text_runlist = [];
        $totruns = 0;
        while ($totruns < $cbRuns) { # counts of BYTES, not runs
            list($rc3, $data3_len, $data3) = $this->book->readRecordParts();
            assert($rc3 == XL_CONTINUE);
            assert($data3_len % 8 == 0);
            for ($pos = 0; $pos < $data3_len; $pos += 8) {
                # <HH4x
                $run = array_values(unpack('v2', substr($data3, $pos, 8)));
                $o->rich_text_runlist[] = $run;
                $totruns += 8;
            }
        }

        # remove trailing entries that point to the end of the string
        while ($o->rich_text_runlist && ($o->rich_text_runlist[count($o->rich_text_runlist)-1][0] == $cchText)) {
            unset($o->rich_text_runlist[count($o->rich_text_runlist)-1]);
        }

        return $o;
    }

    /**
     * @param $data
     * @param $txos MSTxo[]
     */
    public function handle_note($data, $txos)
    {
        $o = new Note();
        $data_len = strlen($data);
        if ($this->biff_version < 80) {
            # <HHH
            list($o->rowx, $o->colx, $expected_bytes) = array_values(unpack("v3", substr($data, 0, 6)));
            $nb = strlen($data) - 6;
            assert($nb <= $expected_bytes);
            $pieces = [substr($data, 6)];
            $expected_bytes -= $nb;
            while ($expected_bytes > 0) {
                list($rc2, $data2_len, $data2) = $this->book->readRecordParts();
                assert($rc2 == XL_NOTE);
                list($dummy_rowx, $nb) = array_values(unpack('v2/x/va', substr($data2, 0, 6)));
                assert($dummy_rowx == 0xffff);
                assert($nb == $data2_len - 6);
                $pieces[] = substr($data2, 6);
                $expected_bytes -= $nb;
            }
            assert($expected_bytes == 0);
            $enc = $this->book->encoding ?: $this->book->deriveEncoding();
            $o->text = join('', $pieces);
            $o->rich_text_runlist = [[0, 0]];
            $o->show = 0;
            $o->row_hidden = 0;
            $o->col_hidden = 0;
            $o->author = '';
            $o->_object_id = Null;
            $this->cell_note_map[$o->rowx][$o->colx] = $o;
            return;
        }
        # Excel 8.0+
        # <HHHH
        list($o->rowx, $o->colx, $option_flags, $o->_object_id) = array_values(unpack("v4", substr($data, 0, 8)));

        $o->show = ($option_flags >> 1) & 1;
        $o->row_hidden = ($option_flags >> 7) & 1;
        $o->col_hidden = ($option_flags >> 8) & 1;
        # XL97 dev kit book says NULL [sic] bytes padding between string count and string data
        # to ensure that string is word-aligned. Appears to be nonsense.
        list($o->author, $endpos) = Helper::unpack_unicode_update_pos($data, 8, 2);
        # There is a random/undefined byte after the author string (not counted in the
        # string length).
        # Issue 4 on github: Google Spreadsheet doesn't write the undefined byte.
        assert(in_array($data_len - $endpos, [0, 1]));
        if (isset($txos[$o->_object_id])) {
            $txo = $txos[$o->_object_id];
            $o->text = $txo->text;
            $o->rich_text_runlist = $txo->rich_text_runlist;
            $this->cell_note_map[$o->rowx][$o->colx] = $o;
        }
    }

    public function string_record_contents($data)
    {

        $debug = $data == "\x0d\x00\x01\x1c\x04\x1e\x04\x21\x04\x1a\x04\x12\x04\x10\x04\x20\x00\x1c\x04\x3e\x04\x41\x04\x3a\x04\x32\x04\x30\x04";

        $bv = $this->biff_version;
        $bk = $this->book;
        $lenlen = (int)($bv >= 30) + 1;
        # <BH
        list($nchars_expected) = array_values(unpack($lenlen-1 ? 'C' : 'v', substr($data, 0, $lenlen)));
        $offset = $lenlen;
        if ($bv < 80) {
            $enc = $bk->encoding ?: $bk->deriveEncoding();
        }
        $nchars_found = 0;
        $result = '';
        $enc = 'latin1';
        while (true) {
            if ($bv >= 80) {
                $flag = ord($data[$offset]) & 1;
                $enc = $flag ? "utf-16le" : "latin1";
                $offset ++;
            }
            $chunk = substr($data, $offset);
            $result .= $chunk;

            # Account for utf16 two byte encoding
            $nchars_found += strlen($chunk) / ($enc == 'latin1' ? 1 : 2);
            if ($nchars_found == $nchars_expected) {
                return mb_convert_encoding($result, 'utf-8', $enc);
            }
            if ($nchars_found > $nchars_expected) {
                throw new XLSParserException(sprintf("STRING/CONTINUE: expected %d chars, found %d", $nchars_expected, $nchars_found));
            }
            list($rc, , $data) = $bk->readRecordParts();
            if ($rc != XL_CONTINUE) {
                throw new XLSParserException(sprintf("Expected CONTINUE record; found record-type 0x%04X", $rc));
            }
            $offset = 0;
        }
    }

    /**
     *
     */
    public function update_cooked_mag_factors()
    {
        # Cached values are used ONLY for the non-active view mode.
        # When the user switches to the non-active view mode,
        # if the cached value for that mode is not valid,
        # Excel pops up a window which says:
        # "The number must be between 10 and 400. Try again by entering a number in this range."
        # When the user hits OK, it drops into the non-active view mode
        # but uses the magn from the active mode.
        # NOTE: definition of "valid" depends on mode ... see below

        if ($this->show_in_page_break_preview) {
            if ($this->scl_mag_factor != Null) { # no SCL record
                $this->cooked_normal_view_mag_factor = 100; # Yes, 100, not 60, NOT a typo
            } else {
                $this->cooked_normal_view_mag_factor = $this->scl_mag_factor;
            }
            $zoom = $this->cached_normal_view_mag_factor;
            if (10 <= $zoom && $zoom <= 400) {
                $zoom = $this->cooked_normal_view_mag_factor;
            }
            $this->cooked_normal_view_mag_factor = $zoom;
        } else {
            # normal view mode
            if ($this->scl_mag_factor == Null) { # no SCL record
                $this->cooked_normal_view_mag_factor = 100;
            } else {
                $this->cooked_normal_view_mag_factor = $this->scl_mag_factor;
            }
            $zoom = $this->cached_normal_view_mag_factor;
            if ($zoom == 0) {
                # VALID. defaults to 60
                $zoom = 60;
            } else if (!(10 <= $zoom && $zoom <= 400)) {
                $zoom = $this->cooked_normal_view_mag_factor;
            }
            $this->cooked_normal_view_mag_factor = $zoom;
        }
    }

    /**
     *
     */
    public function tidy_dimensions()
    {
        if (1 && $this->merged_cells) {
            $nr = $nc = 0;

            foreach ($this->merged_cells as $crange) {
                list($rlo, $rhi, $clo, $chi) = $crange;
                if ($rhi > $nr) $nr = $rhi;
                if ($chi > $nc) $nc = $chi;
            }
            if ($nc > $this->ncols) {
                $this->ncols = $nc;
            }
            if ($nr > $this->nrows) {
                # we put one empty cell at (nr-1,0) to make sure
                # we have the right number of rows. The ragged rows
                # will sort out the rest if needed.
                $this->put_cell($nr - 1, 0, XL_CELL_EMPTY, '', -1);
            }

            if (!$this->ragged_rows) {
                # fix ragged rows
                $ncols = $this->ncols;

                # for rowx in xrange(self.nrows)
                if ($this->first_full_rowx == -2) {
                    $ubound = $this->nrows;
                } else {
                    $ubound = $this->first_full_rowx;
                }
                for ($rowx = 0; $rowx < $ubound; $rowx++) {
                    $trow = $this->cell_types[$rowx];
                    $rlen = strlen($trow);
                    $nextra = $ncols - $rlen;
                    if ($nextra > 0) {
                        $this->cell_values[$rowx] = array_fill(0, $nextra, '');
                        array_pad($trow, count($trow) + $nextra, $this->bt[0]());
                        if ($this->formatting_info) {
                            $trow = array_pad($trow, count($trow) + $nextra, $this->bt[0]());
                        }
                    }
                    $this->cell_types[$rowx] = $trow;
                }
            }
        }
    }

    /**
     * Unragged!
     *
     * @param $rowx
     * @param $colx
     * @param $ctype
     * @param $value
     * @param $xf_index
     */
    public function put_cell($rowx, $colx, $ctype, $value, $xf_index)
    {

        if ($ctype === Null) {
            # we have a number, so look up the cell type
            $ctype = $this->xf_index_to_xl_type_map[$xf_index];
        }

        if (!isset($this->cell_types[$rowx][$colx]) ||
            !isset($this->cell_values[$rowx][$colx]) ||
            ($this->formatting_info && !isset($this->cell_xf_indexes[$rowx][$colx]))
        ) {
            $nr = $rowx + 1;
            $nc = $colx + 1;
            assert(1 <= $nc && $nc <= $this->utter_max_cols);
            assert(1 <= $nr && $nr <= $this->utter_max_rows);

            if ($nc > $this->ncols) {
                $this->ncols = $nc;
                # The row self._first_full_rowx and all subsequent rows
                # are guaranteed to have length == self.ncols. Thus the
                # "fix ragged rows" section of the tidy_dimensions method
                # doesn't need to examine them.
                if ($nr < $this->nrows) {
                    # cell data is not in non-descending row order *AND*
                    # self.ncols has been bumped up.
                    # This very rare case ruins this optmisation.
                    $this->first_full_rowx = -2;
                } elseif ($rowx < $this->first_full_rowx) {
                    $this->first_full_rowx = $rowx;
                }
            }
            if ($nr <= $this->nrows) {
                # New cell is in an existing row, so extend that row (if necessary).
                # Note that nr < self.nrows means that the cell data
                # is not in ascending row order!!
                $trow = $this->cell_types[$rowx];
                $nextra = $this->ncols - count($trow);
                if ($nextra > 0) {
                    $trow = array_pad($trow, count($trow) + $nextra, $this->bt[0]);
                    if ($this->formatting_info) {
                        $this->cell_xf_indexes[$rowx] = array_pad($this->cell_xf_indexes[$rowx], count($this->cell_xf_indexes[$rowx]) + $nextra, $this->bf[0]);
                    }
                    $this->cell_values[$rowx] = array_pad($this->cell_values[$rowx], count($this->cell_values[$rowx]) + $nextra, $this->bf[0]);
                    $this->cells[$rowx] = array_pad($this->cells[$rowx], count($this->cells[$rowx]) + $nextra, new Cell(XL_CELL_EMPTY, '', -1));
                }
            } else {
                $fmt_info = $this->formatting_info;
                $nc = $this->ncols;
                $bt = $this->bt;
                $bf = $this->bf;
                for ($tmp = $this->nrows; $tmp < $nr; $tmp++) {
                    array_push($this->cell_types, array_fill(0, $nc, $bt[0]));
                    array_push($this->cell_values, array_fill(0, $nc, ''));
                    if ($fmt_info) {
                        array_push($this->cell_xf_indexes, array_fill(0, $nc, $bf[0]));
                    }
                    array_push($this->cells, array_fill(0, $nc, new Cell(XL_CELL_EMPTY, '', -1)));
                }
                $this->nrows = $nr;
            }
            $this->cell_types[$rowx][$colx] = $ctype;
            $this->cell_values[$rowx][$colx] = $value;
            if ($this->formatting_info) {
                $this->cell_xf_indexes[$rowx][$colx] = $xf_index;
            }
            $this->cells[$rowx][$colx] = $this->cell($rowx, $colx);
        } else {
            $this->cell_types[$rowx][$colx] = $ctype;
            $this->cell_values[$rowx][$colx] = $value;
            if ($this->formatting_info) {
                $this->cell_xf_indexes[$rowx][$colx] = $xf_index;
            }
            $this->cells[$rowx][$colx] = $this->cell($rowx, $colx);
        }


    }

    public function cell_type($rowx, $colx) {
        return $this->cell_types[$rowx][$colx];
    }

    /**
     * @param $rowx
     * @param $colx
     * @return mixed
     */
    public function cell_value($rowx, $colx) {
        $value = $this->cell($rowx, $colx)->value;
        return $value;
    }

    public function col($colx)
    {
        $cols = [];
        for ($rowx = 0; $rowx < $this->nrows; $rowx++) {
            $cols[] = $this[$rowx][$colx];
        }
        return $cols;
    }

    public function getIterator()
    {
        return new ArrayIterator($this->cells);
    }

    /**
     * Whether a offset exists
     * @link http://php.net/manual/en/arrayaccess.offsetexists.php
     * @param mixed $offset An offset to check for.
     * @return boolean true on success or false on failure.
     * The return value will be casted to boolean if non-boolean was returned.
     */
    public function offsetExists($offset)
    {
        return isset($this->cells[$offset]);
    }

    /**
     * Offset to retrieve
     * @link http://php.net/manual/en/arrayaccess.offsetget.php
     * @param mixed $offset The offset to retrieve.
     * @return mixed Can return all value types.
     */
    public function offsetGet($offset)
    {
        return $this->cells[$offset];
    }

    /**
     * Offset to set
     * @link http://php.net/manual/en/arrayaccess.offsetset.php
     * @param mixed $offset The offset to assign the value to.
     * @param mixed $value The value to set.
     * @return void
     */
    public function offsetSet($offset, $value)
    {
        ;
    }

    /**
     * Offset to unset
     * @link http://php.net/manual/en/arrayaccess.offsetunset.php
     * @param mixed $offset The offset to unset.
     * @return void
     */
    public function offsetUnset($offset)
    {
        ;
    }

}
