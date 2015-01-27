<?php

namespace laacz\XLSParser;

use ArrayAccess;
use ArrayIterator;
use IteratorAggregate;

define('FUN', 0); # unknown
define('FDT', 1); # date
define('FNU', 2); # number
define('FGE', 3); # general
define('FTX', 4); # text

define('oBOOL', 3);
define('oERR',  4);
define('oMSNG', 5); # tMissArg
define('oNUM',  2);
define('oREF', -1);
define('oREL', -2);
define('oSTRG', 1);
define('oUNK',  0);

define('XL_CELL_EMPTY',   0);
define('XL_CELL_TEXT',    1);
define('XL_CELL_NUMBER',  2);
define('XL_CELL_DATE',    3);
define('XL_CELL_BOOLEAN', 4);
define('XL_CELL_ERROR',   5);
define('XL_CELL_BLANK',   6); # for use in debugging, gathering stats, etc


define('NUM_BIG_BLOCK_DEPOT_BLOCKS_POS', 0x2c);
define('SMALL_BLOCK_DEPOT_BLOCK_POS', 0x3c);
define('ROOT_START_BLOCK_POS', 0x30);
define('BIG_BLOCK_SIZE', 0x200);
define('SMALL_BLOCK_SIZE', 0x40);
define('EXTENSION_BLOCK_POS', 0x44);
define('NUM_EXTENSION_BLOCK_POS', 0x48);
define('PROPERTY_STORAGE_BLOCK_SIZE', 0x80);
define('BIG_BLOCK_DEPOT_BLOCKS_POS', 0x4c);
define('SMALL_BLOCK_THRESHOLD', 0x1000);
// property storage offsets
define('SIZE_OF_NAME_POS', 0x40);
define('TYPE_POS', 0x42);
define('START_BLOCK_POS', 0x74);
define('SIZE_POS', 0x78);
define('IDENTIFIER_OLE', pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1));

define('EOCSID',  -2);
define('FREESID', -1);
define('SATSID',  -3);
define('MSATSID', -4);
define('EVILSID', -5);

define('SUPBOOK_UNK',      0);
define('SUPBOOK_INTERNAL', 1);
define('SUPBOOK_EXTERNAL', 2);
define('SUPBOOK_ADDIN',    3);
define('SUPBOOK_DDEOLE',   4);

define('BIFF_FIRST_UNICODE', 80);

define('XL_WORKBOOK_GLOBALS', 0x5);
define('WBKBLOBAL', 0x5);
define('XL_WORKBOOK_GLOBALS_4W', 0x100);
define('XL_WORKSHEET', 0x10);
define('WRKSHEET', 0x10);

define('XL_BOUNDSHEET_WORKSHEET', 0x00);
define('XL_BOUNDSHEET_CHART', 0x02);
define('XL_BOUNDSHEET_VB_MODULE', 0x06);

# XL_RK2 = 0x7e
define('XL_ARRAY', 0x0221);
define('XL_ARRAY2', 0x0021);
define('XL_BLANK', 0x0201);
define('XL_BLANK_B2', 0x01);
define('XL_BOF', 0x809);
define('XL_BOOLERR', 0x205);
define('XL_BOOLERR_B2', 0x5);
define('XL_BOUNDSHEET', 0x85);
define('XL_BUILTINFMTCOUNT', 0x56);
define('XL_CF', 0x01B1);
define('XL_CODEPAGE', 0x42);
define('XL_COLINFO', 0x7D);
define('XL_COLUMNDEFAULT', 0x20);  # BIFF2 only
define('XL_COLWIDTH', 0x24);  # BIFF2 only
define('XL_CONDFMT', 0x01B0);
define('XL_CONTINUE', 0x3c);
define('XL_COUNTRY', 0x8C);
define('XL_DATEMODE', 0x22);
define('XL_DEFAULTROWHEIGHT', 0x0225);
define('XL_DEFCOLWIDTH', 0x55);
define('XL_DIMENSION', 0x200);
define('XL_DIMENSION2', 0x0);
define('XL_EFONT', 0x45);
define('XL_EOF', 0x0a);
define('XL_EXTERNNAME', 0x23);
define('XL_EXTERNSHEET', 0x17);
define('XL_EXTSST', 0xff);
define('XL_FEAT11', 0x872);
define('XL_FILEPASS', 0x2f);
define('XL_FONT', 0x31);
define('XL_FONT_B3B4', 0x231);
define('XL_FORMAT', 0x41e);
define('XL_FORMAT2', 0x1E);  # BIFF2, BIFF3
define('XL_FORMULA', 0x6);
define('XL_FORMULA3', 0x206);
define('XL_FORMULA4', 0x406);
define('XL_GCW', 0xab);
define('XL_HLINK', 0x01B8);
define('XL_QUICKTIP', 0x0800);
define('XL_HORIZONTALPAGEBREAKS', 0x1b);
define('XL_INDEX', 0x20b);
define('XL_INTEGER', 0x2);  # BIFF2 only
define('XL_IXFE', 0x44);  # BIFF2 only
define('XL_LABEL', 0x204);
define('XL_LABEL_B2', 0x04);
define('XL_LABELRANGES', 0x15f);
define('XL_LABELSST', 0xfd);
define('XL_LEFTMARGIN', 0x26);
define('XL_TOPMARGIN', 0x28);
define('XL_RIGHTMARGIN', 0x27);
define('XL_BOTTOMMARGIN', 0x29);
define('XL_HEADER', 0x14);
define('XL_FOOTER', 0x15);
define('XL_HCENTER', 0x83);
define('XL_VCENTER', 0x84);
define('XL_MERGEDCELLS', 0xE5);
define('XL_MSO_DRAWING', 0x00EC);
define('XL_MSO_DRAWING_GROUP', 0x00EB);
define('XL_MSO_DRAWING_SELECTION', 0x00ED);
define('XL_MULRK', 0xbd);
define('XL_MULBLANK', 0xbe);
define('XL_NAME', 0x18);
define('XL_NOTE', 0x1c);
define('XL_NUMBER', 0x203);
define('XL_NUMBER_B2', 0x3);
define('XL_OBJ', 0x5D);
define('XL_PAGESETUP', 0xA1);
define('XL_PALETTE', 0x92);
define('XL_PANE', 0x41);
define('XL_PRINTGRIDLINES', 0x2B);
define('XL_PRINTHEADERS', 0x2A);
define('XL_RK', 0x27e);
define('XL_ROW', 0x208);
define('XL_ROW_B2', 0x08);
define('XL_RSTRING', 0xd6);
define('XL_SCL', 0x00A0);
define('XL_SHEETHDR', 0x8F);  # BIFF4W only
define('XL_SHEETPR', 0x81);
define('XL_SHEETSOFFSET', 0x8E);  # BIFF4W only
define('XL_SHRFMLA', 0x04bc);
define('XL_SST', 0xfc);
define('XL_STANDARDWIDTH', 0x99);
define('XL_STRING', 0x207);
define('XL_STRING_B2', 0x7);
define('XL_STYLE', 0x293);
define('XL_SUPBOOK', 0x1AE);  # aka EXTERNALBOOK in OOo docs
define('XL_TABLEOP', 0x236);
define('XL_TABLEOP2', 0x37);
define('XL_TABLEOP_B2', 0x36);
define('XL_TXO', 0x1b6);
define('XL_UNCALCED', 0x5e);
define('XL_UNKNOWN', 0xffff);
define('XL_VERTICALPAGEBREAKS', 0x1a);
define('XL_WINDOW2', 0x023E);
define('XL_WINDOW2_B2', 0x003E);
define('XL_WRITEACCESS', 0x5C);
define('XL_WSBOOL', XL_SHEETPR);
define('XL_XF', 0xe0);
define('XL_XF2', 0x0043);  # BIFF2 version of XF record
define('XL_XF3', 0x0243);  # BIFF3 version of XF record
define('XL_XF4', 0x0443);  # BIFF4 version of XF record

define('MY_EOF', 0xF00BAAA); # not a 16-bit number

define('USE_FANCY_CD', true);

class XLSParserException extends \Exception {};
class XLSEncryptedException extends \Exception {};
class BookException extends \Exception {};

class Book extends Dumpable implements ArrayAccess, IteratorAggregate
{

    public $base;

    /**
     * @var int
     */
    public $position = 0;

    public $encoding_override = Null;
    public $formatting_info = true;
    public $ragged_rows = false;

    public $sheetoffset = 0;

    /**
     * @var Sheet[]
     */
    public $sheet_list = [];
    public $sheet_names = [];
    public $sheet_visibility = []; # from BOUNDSHEET record
    public $sh_abs_posn = []; # sheet's absolute position in the stream
    public $sharedstrings = [];
    public $rich_text_runlist_map = [];
    public $raw_user_name = False;
    public $sheethdr_count = 0; # BIFF 4W only
    public $builtinfmtcount = -1; # unknown as yet. BIFF 3, 4S, 4W

    public $xfcount = 0;
    public $actualfmtcount = 0; # number of FORMAT records seen so far
    public $xf_index_to_xl_type_map = [0 => XL_CELL_NUMBER];
    public $xf_epilogue_done = 0;

    public $all_sheets_count = 0; # includes macro & VBA sheets
    public $supbook_count = 0;
    public $supbook_locals_inx = Null;
    public $supbook_addins_inx = Null;
    public $all_sheets_map = []; # maps an all_sheets index to a calc-sheets index (or -1)
    public $externsheet_info = [];
    public $externsheet_type_b57 = [];
    public $extnsht_name_from_num = [];
    public $sheet_num_from_name = [];
    public $extnsht_count = 0;
    public $supbook_types = [];
    public $resources_released = 0;
    public $addin_func_names = [];
    public $mem = '';
    public $filestr = '';

    public $stream_len = 0;

    ##
    # The number of worksheets present in the workbook file.
    # This information is available even when no sheets have yet been loaded.
    public $nsheets = 0;

    ##
    # Which date system was in force when this file was last saved.<br />
    #    0 => 1900 system (the Excel for Windows default).<br />
    #    1 => 1904 system (the Excel for Macintosh default).<br />
    public $datemode = 0; # In case it's not specified in the file.

    ##
    # Version of BIFF (Binary Interchange File Format) used to create the file.
    # Latest is 8.0 (represented here as 80), introduced with Excel 97.
    # Earliest supported by this module: 2.0 (represented as 20).
    public $biff_version = 0;

    ##
    # List containing a Name object for each NAME record in the workbook.
    # <br />  -- New in version 0.6.0
    /**
     * @var Name[]
     */
    public $name_obj_list = [];

    ##
    # An integer denoting the character set used for strings in this file.
    # For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.
    # For earlier versions, this is used to derive the appropriate Python encoding
    # to be used to convert to Unicode.
    # Examples: 1252 -> 'cp1252', 10000 -> 'mac_roman'
    public $codepage = Null;

    ##
    # The encoding that was derived from the codepage.
    public $encoding = Null;

    ##
    # A tuple containing the (telephone system) country code for:<br />
    #    [0]: the user-interface setting when the file was created.<br />
    #    [1]: the regional settings.<br />
    # Example: (1, 61) meaning (USA, Australia).
    # This information may give a clue to the correct encoding for an unknown codepage.
    # For a long list of observed values, refer to the OpenOffice.org documentation for
    # the COUNTRY record.
    public $countries = [0, 0];

    ##
    # What (if anything) is recorded as the name of the last user to save the file.
    public $user_name = '';

    ##
    # A list of Font class instances, each corresponding to a FONT record.
    # <br /> -- New in version 0.6.1
    /**
     * @var Font[]
     */
    public $font_list = [];

    ##
    # A list of XF class instances, each corresponding to an XF record.
    # <br /> -- New in version 0.6.1
    /**
     * @var XF[]
     */
    public $xf_list = [];

    ##
    # A list of Format objects, each corresponding to a FORMAT record, in
    # the order that they appear in the input file.
    # It does <i>not</i> contain builtin formats.
    # If you are creating an output file using (for example) pyExcelerator,
    # use this list.
    # The collection to be used for all visual rendering purposes is format_map.
    # <br /> -- New in version 0.6.1
    /**
     * @var Format[]
     */
    public $format_list = [];

    ##
    # The mapping from XF.format_key to Format object.
    # <br /> -- New in version 0.6.1
    public $format_map = [];

    ##
    # This provides access via name to the extended format information for
    # both built-in styles and user-defined styles.<br />
    # It maps <i>name</i> to (<i>built_in</i>, <i>xf_index</i>), where:<br />
    # <i>name</i> is either the name of a user-defined style,
    # or the name of one of the built-in styles. Known built-in names are
    # Normal, RowLevel_1 to RowLevel_7,
    # ColLevel_1 to ColLevel_7, Comma, Currency, Percent, "Comma [0]",
    # "Currency [0]", Hyperlink, and "Followed Hyperlink".<br />
    # <i>built_in</i> 1 = built-in style, 0 = user-defined<br />
    # <i>xf_index</i> is an index into Book.xf_list.<br />
    # References: OOo docs s6.99 (STYLE record); Excel UI Format/Style
    # <br /> -- New in version 0.6.1; since 0.7.4, extracted only if
    # open_workbook(..., formatting_info=True)
    public $style_name_map = [];

    ##
    # This provides definitions for colour indexes. Please refer to the
    # above section "The Palette; Colour Indexes" for an explanation
    # of how colours are represented in Excel.<br />
    # Colour indexes into the palette map into (red, green, blue) tuples.
    # "Magic" indexes e.g. 0x7FFF map to None.
    # <i>colour_map</i> is what you need if you want to render cells on screen or in a PDF
    # file. If you are writing an output XLS file, use <i>palette_record</i>.
    # <br /> -- New in version 0.6.1. Extracted only if open_workbook(..., formatting_info=True)
    public $colour_map = [];
    public $colour_indexes_used = [];

    ##
    # If the user has changed any of the colours in the standard palette, the XLS
    # file will contain a PALETTE record with 56 (16 for Excel 4.0 and earlier)
    # RGB values in it, and this list will be e.g. [(r0, b0, g0), ..., (r55, b55, g55)].
    # Otherwise this list will be empty. This is what you need if you are
    # writing an output XLS file. If you want to render cells on screen or in a PDF
    # file, use colour_map.
    # <br /> -- New in version 0.6.1. Extracted only if open_workbook(..., formatting_info=True)
    public $palette_record = [];

    public function __construct($data, Array $config = [])
    {

        foreach ($config as $k=>$v) {
            $this->$k = $v;
        }

        foreach (Defs::$fmt_code_ranges as $range) {
            list($lo, $hi, $ty) = $range;
            for ($x = $lo; $x <= $hi; $x++) {
                $this->std_format_code_types[$x] = $ty;
            }
        }

        $this->mem = $data;
        $this->loadBiff28();
        $biff_version = $this->readBOF(XL_WORKBOOK_GLOBALS);
        if (!$biff_version) {
            throw new XLSParserException("Can't determine file's BIFF version");
        }
        if (!in_array($biff_version, Defs::$supported_versions)) {
            throw new XLSParserException(sprintf("BIFF version %s is not supported", Defs::$biff_text_from_num[$biff_version]));
        }

        $this->biff_version = $biff_version;

        if ($biff_version <= 40) {
            $this->fake_globals_get_sheet();
        } elseif ($biff_version == 45) {
            $this->parseGlobals();
        } else {
            $this->parseGlobals();
            $this->sheet_list = array_fill(0, count($this->sheet_names), null);
            $this->readSheets();
        }

        $this->nsheets = count($this->sheet_list);

        $this->initialiseBook();

    }

    public function initialiseColourMap()
    {
        $this->colour_map = [];
        $this->colour_indexes_used = [];
        if (!$this->formatting_info) {
            return;
        }
        # Add the 8 invariant colours
        for ($i = 0; $i < 8; $i++) {
            $this->colour_map[$i] = Defs::$excel_default_palette_b8[$i];
        }
        # Add the default palette depending on the version
        $dpal = Defs::${Defs::$default_palette[$this->biff_version]};
        $ndpal = count($dpal);
        for ($i = 0; $i < $ndpal; $i++) {
            $this->colour_map[$i+8] = $dpal[$i];
        }
        # Add the specials -- None means the RGB value is not known
        # System window text colour for border lines
        $this->colour_map[$ndpal+8] = Null;
        # System window background colour for pattern background
        $this->colour_map[$ndpal+8+1] = Null; #
        $this->colour_map[$ndpal+8+1] = Null; #
        $this->colour_map[0x51] = Null; # System ToolTip text colour (used in note objects)
        $this->colour_map[0x7FFF] = Null; # 32767, system window text colour for fonts

    }

    public function readSheets()
    {
        Helper::debug("GET_SHEETS: %s (%s)", join(', ', $this->sheet_names), join(', ', $this->sh_abs_posn));
        for ($sheetno = 0; $sheetno < count($this->sheet_names); $sheetno++) {
            if (!isset($this->sheet_list[$sheetno])) {
                Helper::log("GET_SHEETS: sheetno = %d, %s (%s)", $sheetno, join(', ', $this->sheet_names), join(', ', $this->sh_abs_posn));
                $this->readSheet($sheetno);
            }
        }
    }

    public function getSheetByName($sheet_name)
    {
        if (($sheetx = array_search($sheet_name, $this->sheet_names, true)) !== false) {
            return $this->sheet_list[$sheetx];
        }
        throw new BookException("No such sheet: $sheet_name");
    }

    public function getSheetByIndex($sheetx) {
        return isset($this->sheet_list[$sheetx]) ? $this->sheet_list[$sheetx] : $this->readSheet($sheetx);
    }

    /**
     * @param $sh_number
     * @param bool $update_pos
     * @return Sheet
     * @throws XLSParserException
     */
    public function readSheet($sh_number, $update_pos = true)
    {
        if ($update_pos) {
            $this->position = $this->sh_abs_posn[$sh_number];
        }
        # Advance BOF
        $this->readBOF(XL_WORKSHEET);
        # assert biff_version == $this->biff_version ### FAILS
        # Have an example where book is v7 but sheet reports v8!!!
        # It appears to work OK if the sheet version is ignored.
        # Confirmed by Daniel Rentz: happens when Excel does "save as"
        # creating an old version file; ignore version details on sheet BOF.
        $sh = new Sheet($this, $this->position, $this->sheet_names[$sh_number], $sh_number);
        $sh->read($this);
        $this->sheet_list[$sh_number] = $sh;
        return $sh;
    }

    public function initialiseBook()
    {
        $this->initialiseColourMap();
        $this->xf_epilogue_done = 0;
    }

    public function fake_globals_get_sheet()
    {
        $this->initialiseBook();
        $fake_sheet_name = 'Sheet 1';
        $this->sheet_names = [$fake_sheet_name];
        $this->sh_abs_posn = [0];
        $this->sheet_visibility = [0];
        $this->sheet_list[] = Null;
        $this->readSheets();
    }

    public function parseGlobals()
    {
        $this->initialiseBook();
        $datalen = strlen($this->mem);
        $rc = false;
        while ($rc != XL_EOF && ($this->position < $datalen)) {
            list($rc, $length, $data) = $this->readRecordParts();
            Helper::debug("parse_globals: record code is 0x%04x (%s)", $rc, Helper::explain_const($rc, 'XL_') ?: 'N/A');
            switch ($rc) {
                case XL_SST:
                    $this->handleSST($data);
                    break;
                case XL_FONT:
                case XL_FONT_B3B4:
                    $this->handleFont($data);
                    break;
                case XL_FORMAT: # XL_FORMAT2 is BIFF <= 3.0, can't appear in globals
                    $this->handleFormat($data);
                    break;
                case XL_XF:
                    $this->handleXF($data);
                    break;
                case  XL_BOUNDSHEET:
                    $this->handleBoundSheet($data);
                    break;
                case XL_DATEMODE:
                    $this->handleDateMode($data);
                    break;
                case XL_CODEPAGE:
                    $this->handleCodePage($data);
                    break;
                case XL_COUNTRY:
                    $this->handleCountry($data);
                    break;
                case XL_EXTERNNAME:
                    $this->handle_externname($data);
                    break;
                case XL_EXTERNSHEET:
                    $this->handleExternSheet($data);
                    break;
                case XL_FILEPASS:
                    $this->handle_filepass($data);
                    break;
                case XL_WRITEACCESS:
                    $this->handleWriteAccess($data);
                    break;
                case XL_SHEETSOFFSET:
                    $this->handle_sheetsoffset($data);
                    break;
                case XL_SHEETHDR:
                    $this->handle_sheethdr($data);
                    break;
                case XL_SUPBOOK:
                    $this->handleSupbook($data);
                    break;
                case XL_NAME:
                    $this->handleName($data);
                    break;
                case XL_PALETTE:
                    $this->handlePalette($data);
                    break;
                case XL_STYLE:
                    $this->handleStyle($data);
                    break;
                case XL_EOF:
                    $this->XFEpilogue();
                    $this->namesEpilogue();
                    $this->paletteEpilogue();
                    if (!$this->encoding) {
                        $this->deriveEncoding();
                    }
                    if ($this->biff_version == 45) {
                        Helper::debug("global EOF: position %d", $this->position);
                    }
                    return;
                    break;
                default:
                    if ($rc & 0xff == 9) {
                        Helper::debug("*** Unexpected BOF at posn %d: 0x%04x len=%d data=%s\n",
                            $this->position - $length - 4, $rc, $length, Helper::as_hex($data));
                    }

            }
        }
    }

    public function handle_filepass($data)
    {
        throw new XLSEncryptedException();
    }

    public function handle_externname($data)
    {
        if ($this->biff_version >= 80) {
            # <HI
            list($option_flags, $other_info) = array_values(unpack('va/Ib', substr($data, 0, 6)));
            $pos = 6;
            list($name, $pos) = Helper::unpack_unicode_update_pos($data, $pos, 1);
            $extra = substr($data, $pos);
            if ($this->supbook_types[-1] == SUPBOOK_ADDIN) {
                $this->addin_func_names[]= $name;
            }
        }
    }

    public function handleName($data)
    {
        $bv = $this->biff_version;
        if ($bv < 50) {
            return;
        }
        $this->deriveEncoding();
        # <HBBHHH4B
        list($option_flags, $kb_shortcut, $name_len, $fmla_len, $extsht_index, $sheet_index, $menu_text_len, $description_text_len, $help_topic_text_len, $status_bar_text_len) = array_values(unpack('va/C2b/v3c/C4d', substr($data, 0, 14)));
        $nobj = new Name();
        $nobj->book = $this; ### CIRCUAL ###
        $name_index = count($this->name_obj_list);
        $this->name_obj_list[] = $nobj;
        # Not used anywhere in code. Commenting out.
        # $nobj->option_flags = $option_flags;
        $attrs = [
            ['hidden', 1, 0],
            ['func', 2, 1],
            ['vbasic', 4, 2],
            ['macro', 8, 3],
            ['complex', 0x10, 4],
            ['builtin', 0x20, 5],
            ['funcgroup', 0xFC0, 6],
            ['binary', 0x1000, 12],
        ];
        foreach ($attrs as $row) {
            list($attr, $mask, $nshift) = $row;
            $nobj->$attr = ($option_flags & $mask) >> $nshift;
        }
        $macro_flag = $nobj->macro ? 'M' : ' ';
        if ($bv < 80) {
            list($internal_name, $pos) = Helper::unpack_string_update_pos($data, 14, $this->encoding, $name_len);
        } else {
            list($internal_name, $pos) = Helper::unpack_unicode_update_pos($data, 14, 2, $name_len);
        }
        $nobj->extn_sheet_num = $extsht_index;
        $nobj->excel_sheet_index = $sheet_index;
        $nobj->scope = Null; # patched up in the names_epilogue() method
        Helper::debug("NAME[%d]:%s oflags=%d, name_len=%d, fmla_len=%d, extsht_index=%d, sheet_index=%d, name=%s, %d",
            $name_index, $macro_flag, $option_flags, $name_len,
            $fmla_len, $extsht_index, $sheet_index, $internal_name ,strlen($internal_name));
        $name = $internal_name;
        if ($nobj->builtin) {
            $name = isset(Defs::$builtin_name_from_code[$name]) ? Defs::$builtin_name_from_code[$name] : "??Unknown??";
            Helper::debug("    builtin: %s", $name);
        }
        $nobj->name = $name;
        $nobj->raw_formula = substr($data, $pos);
        $nobj->basic_formula_len = $fmla_len;
        $nobj->evaluated = 0;
        Helper::debug($nobj->dump("--- handle_name: name[$name_index] ---", '-------------------'));
    }

    public function handleWriteAccess($data)
    {
        if ($this->biff_version < 80) {
            if (!$this->encoding) {
                $this->raw_user_name = true;
                $this->user_name = $data;
                return;
            }
            $strg = Helper::unpack_string($data, 0, $this->encoding, 1);
        } else {
            $strg = Helper::unpack_unicode($data, 0, 2);
        }
        Helper::debug("WRITEACCESS: %d bytes; raw=%s %s", strlen($data), (int)$this->raw_user_name, $strg);
        $strg = rtrim($strg);
        $this->user_name = $strg;
    }

    public function handleFont($data)
    {
        IF (!$this->formatting_info) {
            return;
        }
        if (!$this->encoding) {
            $this->deriveEncoding();
        }
        $bv = $this->biff_version;
        $k = count($this->font_list);
        if ($k == 4) {
            $f = new Font();
            $f->name = "Dummy font";
            $f->font_index = $k;
            $this->font_list[] = $f;
            $k ++;
        }
        $f = new Font();
        $f->font_index = $k;
        $this->font_list[] = $f;
        if ($bv >= 50) {
            # <HHHHHBBB
            list($f->height, $option_flags, $f->colour_index, $f->weight, $f->escapement, $f->underline_type, $f->family, $f->character_set) = array_values(unpack('v5a/C3b', substr($data, 0, 13)));
            $f->bold = $option_flags & 1;
            $f->italic = ($option_flags & 2) >> 1;
            $f->underlined = ($option_flags & 4) >> 2;
            $f->struck_out = ($option_flags & 8) >> 3;
            $f->outline = ($option_flags & 16) >> 4;
            $f->shadow = ($option_flags & 32) >> 5;
            if ($bv >= 80) {
                $f->name = Helper::unpack_unicode($data, 14, 1);
            } else {
                $f->name = Helper::unpack_string($data, 14, $this->encoding, 1);
            }
        } elseif ($bv >= 30) {
            # <HHH
            list($f->height, $option_flags, $f->colour_index) = array_values(unpack('v3', substr($data, 0, 6)));
            $f->bold = $option_flags & 1;
            $f->italic = ($option_flags & 2) >> 1;
            $f->underlined = ($option_flags & 4) >> 2;
            $f->struck_out = ($option_flags & 8) >> 3;
            $f->outline = ($option_flags & 16) >> 4;
            $f->shadow = ($option_flags & 32) >> 5;
            $f->name = Helper::unpack_string($data, 6, $this->encoding, 1);
            # Now cook up the remaining attributes ...
            $f->weight = $f->bold == 0 ? 400 : 700;
            $f->escapement = 0; # None
            $f->underline_type = $f->underlined; # None or Single
            $f->family = 0; # Unknown / don't care
            $f->character_set = 1; # System default (0 means "ANSI Latin")
        } else { # BIFF2
            # <HH
            list($f->height, $option_flags) = array_values(unpack('v2', substr($data, 0, 4)));
            $f->colour_index = 0x7FFF; # "system window text colour"
            $f->bold = $option_flags & 1;
            $f->italic = ($option_flags & 2) >> 1;
            $f->underlined = ($option_flags & 4) >> 2;
            $f->struck_out = ($option_flags & 8) >> 3;
            $f->outline = 0;
            $f->shadow = 0;
            $f->name = Helper::unpack_string($data, 4, $this->encoding, 1);
            # Now cook up the remaining attributes ...
            $f->weight = $f->bold == 0 ? 400 : 700;;
            $f->escapement = 0; # None
            $f->underline_type = $f->underlined; # None or Single
            $f->family = 0; # Unknown / don't care
            $f->character_set = 1; # System default (0 means "ANSI Latin")

        }

    }
    // @TODO: From Formatting
//    public function handle_efont($data) {}

    public function handleFormat($data, $rectype = XL_FORMAT)
    {
        $bv = $this->biff_version;
        if ($rectype == XL_FORMAT2) {
            $bv = min($bv, 30);
        }
        if (!$this->encoding) {
            $this->deriveEncoding();
        }
        $strpos = 2;
        if ($bv >= 50) {
            # <H
            list($fmtkey) = array_values(unpack('v', substr($data, 0, 2)));
        } else {
            $fmtkey = $this->actualfmtcount;
            if ($bv <= 30) {
                $strpos = 0;
            }
        }

        $this->actualfmtcount ++;

        if ($bv >= 80) {
            $unistrg = Helper::unpack_unicode($data, 2);
        } else {
            $unistrg = Helper::unpack_string($data, $strpos, $this->encoding, 1);
        }
        $is_date_s = $this->isDateFormatString($unistrg);
        $ty = $is_date_s ? FDT : FGE;
        if ($fmtkey < 163 && $fmtkey > 50) {
            # user_defined if fmtkey > 163
            # N.B. Gnumeric incorrectly starts these at 50 instead of 164 :-(
            # if earlier than BIFF 5, standard info is useless
            # Comment out unused variables.
            # $std_ty = isset($this->std_format_code_types[$fmtkey]) ? $this->std_format_code_types[$fmtkey] : FUN;
            # $is_date_c = $std_ty = FDT;
        }
        $fmtobj = new Format($fmtkey, $ty, $unistrg);
        $this->format_map[$fmtkey] = $fmtobj;
        $this->format_list[] = $fmtobj;

    }

    public function handlePalette($data)
    {
        if (!$this->formatting_info) return;
        # <H
        list($n_colours) = array_values(unpack('v', substr($data, 0, 2)));
        # $expected_n_colours = $this->biff_version >= 50 ? 56 : 16;

        # <'<xx%di' % n_colours # use i to avoid long integers
        $fmt = "x2/i{$n_colours}";
        $expected_size = 4 * $n_colours + 2;
        $actual_size = strlen($data);
        $tolerance = 4;
        if (!($expected_size <= $actual_size && $actual_size <= $expected_size + $tolerance)) {
            throw new XLSParserException(sprintf('PALETTE record: expected size %d, actual size %d', $expected_size, $actual_size));
        }
        $colours = array_values(unpack($fmt, substr($data, 0, $expected_size)));

        assert($this->palette_record == []);
        # a colour will be 0xbbggrr
        # IOW, red is at the little end
        for ($i = 0; $i < $n_colours; $i++) {
            $c = $colours[$i];
            $red   = $c        & 0xff;
            $green = ($c >>  8) & 0xff;
            $blue  = ($c >> 16) & 0xff;
            # $old_rgb = $this->colour_map[8 + $i];
            $new_rgb = [$red, $green, $blue];

            $this->palette_record[] = $new_rgb;
        }

    }

    public function paletteEpilogue()
    {
        # Check colour indexes in fonts etc.
        # This must be done here as FONT records
        # come *before* the PALETTE record :-(
        foreach ($this->font_list as $font) {
            if ($font->font_index == 4) { # the missing font record
                continue;
            }
            $cx = $font->colour_index;
            if ($cx == 0x7fff) { # system window text colour
               continue;
            }
            if (isset($this->colour_map[$cx])) {
                $this->colour_indexes_used[$cx] = 1;
            }
        }

    }

    public function handleStyle($data)
    {
        if (!$this->formatting_info) return;
        $bv = $this->biff_version;
        # <HBB
        list($flag_and_xfx, $built_in_id, $level) = array_values(unpack('va/C2b', substr($data, 0, 4)));
        $xf_index = $flag_and_xfx & 0x0fff;
        if ($data == "\x00\x00\x00\x00" && !isset($this->style_name_map['Normal'])) {
            # Erroneous record (doesn't have built-in bit set).
            # Example file supplied by Jeff Bell.
            $built_in = 1;
            # $built_in_id = 0;
            $xf_index = 0;
            $name = "Normal";
            # $level = 255;
        } elseif ($flag_and_xfx & 0x8000) {
            # built-in style
            $built_in = 1;
            $name = Defs::$built_in_style_names[$built_in_id];
            if ($built_in_id >= 1 && $built_in_id <= 2) {
                $name .= (string)($level + 1);
            }
        } else {
            # user-defined style
            $built_in = 0;
            # $built_in_id = 0;
            # $level = 0;
            if ($bv >= 80) {
                $name = Helper::unpack_unicode($data, 2, 2);
            } else {
                $name = Helper::unpack_string($data, 2, $this->encoding, 1);
            }
        }
        $this->style_name_map[$name] = [$built_in, $xf_index];
    }

    public function isDateFormatString($fmt)
    {
        # Heuristics:
        # Ignore "text" and [stuff in square brackets (aarrgghh -- see below)].
        # Handle backslashed-escaped chars properly.
        # E.g. hh\hmm\mss\s should produce a display like 23h59m59s
        # Date formats have one or more of ymdhs (caseless) in them.
        # Numeric formats have # and 0.
        # N.B. u'General"."' hence get rid of "text" first.
        # TODO: Find where formats are interpreted in Gnumeric
        # TODO: u'[h]\\ \\h\\o\\u\\r\\s' ([h] means don't care about hours > 23)
        $state = 0;
        $s = '';
        $non_date_formats = [
          '0.00E+00',
          '##0.0E+0',
          'General',
          'GENERAL', # OOo Calc 1.1.4 does this.
          'general', # pyExcelerator 0.6.3 does this.
          '@',
        ];

        foreach (str_split($fmt) as $c) {
            if ($state == 0) {
                if ($c == '"') {
                    $state = 1;
                } elseif (in_array($c, str_split('\_*'))) {
                    $state = 2;
                } elseif (in_array($c, str_split('$-+/()'))) {
                    ;
                } else {
                    $s .= $c;
                }
            } elseif ($state == 1) {
                if ($c == '"') {
                    $state = 0;
                }
            } elseif ($state == 2) {
                # Ignore char after backslash, underscore or asterisk
                $state = 0;
            }
            assert($state >= 0 && $state <= 2);
        }
        $s = preg_replace('|\[[^]]*\]|', '', $s);
        if (in_array($s, $non_date_formats)) {
            return false;
        }

        # $separator = ";";
        $date_count = $num_count = 0;
        foreach (str_split($s) as $c) {
            if (in_array($c, str_split("ymdhsYMDHS"))) {
                $date_count += 5;
            } elseif (in_array($c, str_split("0#?"))) {
                $num_count += 5;
            # } elseif ($c == $separator) {
            #     $got_sep = 1;
            }
        }
        #print num_count, date_count, repr(fmt);
        if ($date_count && !$num_count) {
            return true;
        }
        if ($num_count && !$date_count) {
            return false;
        }
        return $date_count > $num_count;
    }

    public function handleXF($data) {

        ### self is a Book instance
        $bv = $this->biff_version;
        $xf = new XF();
        $xf->alignment = new XFAlignment();
        $xf->alignment->indent_level = 0;
        $xf->alignment->shrink_to_fit = 0;
        $xf->alignment->text_direction = 0;
        $xf->border = new XFBorder();
        $xf->border->diag_up = 0;
        $xf->border->diag_down = 0;
        $xf->border->diag_colour_index = 0;
        $xf->border->diag_line_style = 0; # no line
        $xf->background = new XFBackground();
        $xf->protection = new XFProtection();

        # fill in the known standard formats
        if ($bv >= 50 and !$this->xfcount) {
            # i.e. do this once before we process the first XF record
            $this->fillInStandardFormats();
        }

        if ($bv >= 80) {
            # <HHHBBBBIiH
            $unpack_fmt = "v3a/C4b/Vc/id/ve";
            list ($xf->font_index, $xf->format_key, $pkd_type_par, $pkd_align1, $xf->alignment_rotation, $pkd_align2, $pkd_used, $pkd_brdbkg1, $pkd_brdbkg2, $pkd_brdbkg3) = array_values(unpack($unpack_fmt, substr($data, 0, 20)));

            Helper::upkbits($xf->protection, $pkd_type_par, [
                [0, 0x01, 'cell_locked'],
                [1, 0x02, 'formula_hidden'],
            ]);
            Helper::upkbits($xf, $pkd_type_par, [
                [2, 0x0004, 'is_style'],
                # Following is not in OOo docs, but is mentioned
                # in Gnumeric source and also in (deep breath)
                # org.apache.poi.hssf.record.ExtendedFormatRecord.java
                [3, 0x0008, 'lotus_123_prefix'], # Meaning is not known.
                [4, 0xFFF0, 'parent_style_index'],
            ]);
            Helper::upkbits($xf->alignment, $pkd_align1, [
                [0, 0x07, 'hor_align'],
                [3, 0x08, 'text_wrapped'],
                [4, 0x70, 'vert_align'],
            ]);
            Helper::upkbits($xf->alignment, $pkd_align2, [
                [0, 0x0f, 'indent_level'],
                [4, 0x10, 'shrink_to_fit'],
                [6, 0xC0, 'text_direction'],
            ]);

            $reg = $pkd_used >> 2;

            foreach (['format', 'font', 'alignment', 'border', 'background', 'protection'] as $attr_stem) {
                $attr = "_" . $attr_stem . "_flag";
                $xf->$attr = $reg & 1;
                $reg >>= 1;
            }

            Helper::upkbitsL($xf->border, $pkd_brdbkg1, [
                [0,  0x0000000f, 'left_line_style'],
                [4,  0x000000f0, 'right_line_style'],
                [8,  0x00000f00, 'top_line_style'],
                [12, 0x0000f000, 'bottom_line_style'],
                [16, 0x007f0000, 'left_colour_index'],
                [23, 0x3f800000, 'right_colour_index'],
                [30, 0x40000000, 'diag_down'],
                [31, 0x80000000, 'diag_up'],
            ]);
            Helper::upkbits($xf->border, $pkd_brdbkg2, [
                [0,  0x0000007F, 'top_colour_index'],
                [7,  0x00003F80, 'bottom_colour_index'],
                [14, 0x001FC000, 'diag_colour_index'],
                [21, 0x01E00000, 'diag_line_style'],
            ]);
            Helper::upkbitsL($xf->background, $pkd_brdbkg2, [
                [26, 0xFC000000, 'fill_pattern'],
            ]);
            Helper::upkbits($xf->background, $pkd_brdbkg3, [
                [0, 0x007F, 'pattern_colour_index'],
                [7, 0x3F80, 'background_colour_index'],
            ]);
        } elseif ($bv >= 50) {
            # <HHHBBIi
            $unpack_fmt = "v3a/C2b/Vc/id";
            list ($xf->font_index, $xf->format_key, $pkd_type_par, $pkd_align1, $pkd_orient_used, $pkd_brdbkg1, $pkd_brdbkg2) = array_values(unpack($unpack_fmt, substr($data, 0, 16)));

            Helper::upkbits($xf->protection, $pkd_type_par, [
                [0, 0x01, 'cell_locked'],
                [1, 0x02, 'formula_hidden'],
            ]);
            Helper::upkbits($xf, $pkd_type_par, [
                [2, 0x0004, 'is_style'],
                [3, 0x0008, 'lotus_123_prefix'], # Meaning is not known.
                [4, 0xFFF0, 'parent_style_index'],
            ]);
            Helper::upkbits($xf->alignment, $pkd_align1, [
                [0, 0x07, 'hor_align'],
                [3, 0x08, 'text_wrapped'],
                [4, 0x70, 'vert_align'],
            ]);

            $orientation = $pkd_orient_used & 0x03;
            $rotations = [0, 255, 90, 180];
            $xf->alignment->rotation = $rotations[$orientation];
            $reg = $pkd_orient_used >> 2;
            foreach (['format', 'font', 'alignment', 'border', 'background', 'protection'] as $attr_stem) {
                $attr = "_{$attr_stem}_flag";
                $xf->$attr = $reg & 1;
                $reg >>= 1;
            }
            Helper::upkbitsL($xf->background, $pkd_brdbkg1, [
                [ 0, 0x0000007F, 'pattern_colour_index'],
                [ 7, 0x00003F80, 'background_colour_index'],
                [16, 0x003F0000, 'fill_pattern'],
            ]);
            Helper::upkbitsL($xf->border, $pkd_brdbkg1, [
                [22, 0x01C00000,  'bottom_line_style'],
                [25, 0xFE000000, 'bottom_colour_index'],
            ]);
            Helper::upkbits($xf->border, $pkd_brdbkg2, [
                [ 0, 0x00000007, 'top_line_style'],
                [ 3, 0x00000038, 'left_line_style'],
                [ 6, 0x000001C0, 'right_line_style'],
                [ 9, 0x0000FE00, 'top_colour_index'],
                [16, 0x007F0000, 'left_colour_index'],
                [23, 0x3F800000, 'right_colour_index'],
            ]);
            // @TODO
        } elseif ($bv >= 40) {
            # <BBHBBHI
            list($xf->font_index, $xf->format_key, $pkd_type_par, $pkd_align_orient, $pkd_used, $pkd_bkg_34, $pkd_brd_34) = array_values(unpack('C2a/vb/C2c/vd/Ve', substr($data, 0, 12)));

            Helper::upkbits($xf->protection, $pkd_type_par, [
                [0, 0x01, 'cell_locked'],
                [1, 0x02, 'formula_hidden'],
            ]);
            Helper::upkbits($xf, $pkd_type_par, [
                [2, 0x0004, 'is_style'],
                [3, 0x0008, 'lotus_123_prefix'], # Meaning is not known.
                [4, 0xFFF0, 'parent_style_index'],
            ]);
            Helper::upkbits($xf->$alignment, $pkd_align_orient, [
                [0, 0x07, 'hor_align'],
                [3, 0x08, 'text_wrapped'],
                [4, 0x30, 'vert_align'],
            ]);
            $orientation = ($pkd_align_orient & 0xC0) >> 6;
            $rotations = [0, 255, 90, 180];
            $xf->alignment->rotation = $rotations[$orientation];
            $reg = $pkd_used >> 2;
            foreach (['format', 'font', 'alignment', 'border', 'background', 'protection'] as $attr_stem) {
                $attr = "_{$attr_stem}_flag";
                $xf->$attr = $reg & 1;
                $reg >>= 1;
            }
            Helper::upkbits($xf->background, $pkd_bkg_34, [
                [ 0, 0x003F, 'fill_pattern'],
                [ 6, 0x07C0, 'pattern_colour_index'],
                [11, 0xF800, 'background_colour_index'],
            ]);
            Helper::upkbitsL($xf->border, $pkd_brd_34, [
                [ 0, 0x00000007,  'top_line_style'],
                [ 3, 0x000000F8,  'top_colour_index'],
                [ 8, 0x00000700,  'left_line_style'],
                [11, 0x0000F800,  'left_colour_index'],
                [16, 0x00070000,  'bottom_line_style'],
                [19, 0x00F80000,  'bottom_colour_index'],
                [24, 0x07000000,  'right_line_style'],
                [27, 0xF8000000, 'right_colour_index'],
            ]);

        } elseif ($bv == 30) {
            # <BBBBHHI
            list($xf->font_index, $xf->format_key, $pkd_type_prot, $pkd_used, $pkd_align_par, $pkd_bkg_34, $pkd_brd_34) = array_values(unpack('C4a/v2b/Vc', substr($data, 0, 12)));

            Helper::upkbits($xf->protection, $pkd_type_prot, [
                [0, 0x01, 'cell_locked'],
                [1, 0x02, 'formula_hidden'],
            ]);
            Helper::upkbits($xf, $pkd_type_prot, [
                [2, 0x0004, 'is_style'],
                [3, 0x0008, 'lotus_123_prefix'], # Meaning is not known.
            ]);
            Helper::upkbits($xf->alignment, $pkd_align_par, [
                [0, 0x07, 'hor_align'],
                [3, 0x08, 'text_wrapped'],
            ]);
            Helper::upkbits($xf, $pkd_align_par, [
                [4, 0xFFF0, 'parent_style_index'],
            ]);
            $reg = $pkd_used >> 2;

            foreach (['format', 'font', 'alignment', 'border', 'background', 'protection'] as $attr_stem) {
                $attr = "_{$attr_stem}_flag";
                $xf->$attr = $reg & 1;
                $reg >>= 1;
            }
            Helper::upkbits($xf->background, $pkd_bkg_34, [
                [ 0, 0x003F, 'fill_pattern'],
                [ 6, 0x07C0, 'pattern_colour_index'],
                [11, 0xF800, 'background_colour_index'],
            ]);
            Helper::upkbitsL($xf->border, $pkd_brd_34, [
                [ 0, 0x00000007,  'top_line_style'],
                [ 3, 0x000000F8,  'top_colour_index'],
                [ 8, 0x00000700,  'left_line_style'],
                [11, 0x0000F800,  'left_colour_index'],
                [16, 0x00070000,  'bottom_line_style'],
                [19, 0x00F80000,  'bottom_colour_index'],
                [24, 0x07000000,  'right_line_style'],
                [27, 0xF8000000, 'right_colour_index'],
            ]);
            $xf->alignment->vert_align = 2; # bottom
            $xf->alignment->rotation = 0;
        } elseif ($bv == 21) {
            #### Warning: incomplete treatment; formatting_info not fully supported.
            #### Probably need to offset incoming BIFF2 XF[n] to BIFF8-like XF[n+16],
            #### and create XF[0:16] like the standard ones in BIFF8
            #### *AND* add 16 to all XF references in cell records :-(
            # <BxBB
            list($xf->font_index, /* ignore second */ ,$format_etc, $halign_etc) = array_values(unpack('C4', $data));
            $xf->format_key = $format_etc & 0x3F;
            Helper::upkbits($xf->protection, $format_etc, [
                [6, 0x40, 'cell_locked'],
                [7, 0x80, 'formula_hidden'],
            ]);
            Helper::upkbits($xf->alignment, $halign_etc, [
                [0, 0x07, 'hor_align'],
            ]);

            foreach ([[0x08, 'left'], [0x10, 'right'], [0x20, 'top'], [0x40, 'bottom']] as $row) {
                list($mask, $side) = $row;
                if ($halign_etc && $mask) {
                    $colour_index = 8; # black
                    $line_style = 1; # thin
                } else {
                    $colour_index = 0; # none
                    $line_style = 0; # none
                }
                $xf->border->{"{$side}_colour_index"} = $colour_index;
                $xf->border->{"{$side}_line_style"} = $line_style;
            }
            $bg = $xf->background;
            if ($halign_etc & 0x80) {
                $bg->fill_pattern = 17;
            } else {
                $bg->fill_pattern = 0;
            }
            $bg->background_colour_index = 9; # white
            $bg->pattern_colour_index = 8; # black
            $xf->parent_style_index = 0; # ??????????
            $xf->alignment->vert_align = 2; # bottom
            $xf->alignment->rotation = 0;

            foreach (['format', 'font', 'alignment', 'border', 'background', 'protection'] as $attr_stem) {
                $attr = "_{$attr_stem}_flag";
                $xf->$attr = 1;
            }
        } else {
            throw new XLSParserException("programmer stuff-up: bv=$bv");
        }

        $xf->xf_index = count($this->xf_list);

        $this->xf_list[] = $xf;
        $this->xfcount ++;

        $fmt = $cellty = null;

        if (isset($this->format_map[$xf->format_key])) {
            $fmt = $this->format_map[$xf->format_key];
            $cellty = isset(Defs::$cellty_from_fmtty[$fmt->type]) ? Defs::$cellty_from_fmtty[$fmt->type] : false;
        }
        if (!$fmt || !$cellty) {
            $cellty = XL_CELL_NUMBER;
        }

        $this->xf_index_to_xl_type_map[$xf->xf_index] = $cellty;

        # Now for some assertions ...
        if ($this->formatting_info) {
            $this->checkColourIndexesInObj($xf, $xf->xf_index);
        }

        if (!isset($this->format_map[$xf->format_key])) {
            $xf->format_key = 0;
        }

    }

    public function checkColourIndexesInObj($obj, $orig_index) {
        $alist = get_object_vars($obj);
        foreach ($alist as $attr => $nobj) {
            if (is_object($nobj)) {
                $this->checkColourIndexesInObj($nobj, $orig_index);
            } elseif (isset($nobj->colour_index)) {
                if (isset($this->colour_map[$nobj])) {
                    $this->colour_indexes_used[$nobj] = 1;
                    continue;
                }
            }
        }
    }

    public function fillInStandardFormats()
    {
        foreach (array_keys($this->std_format_code_types) as $x) {
            if (!isset($this->format_map[$x])) {
                $ty = $this->std_format_code_types[$x];
                $fmt_str = isset(Defs::$std_format_strings[$x]) ? Defs::$std_format_strings[$x] : Null;
                $fmtobj = new Format($x, $ty, $fmt_str);
                $this->format_map[$x] = $fmtobj;
            }
        }
    }

    public function XFEpilogue()
    {
        # self is a Book instance
        $this->xf_epilogue_done = 1;
        $num_xfs = count($this->xf_list);

        for ($xfx = 0; $xfx < $num_xfs; $xfx++) {
            $xf = $this->xf_list[$xfx];
            if (!isset($this->format_map[$xf->format_key])) {
                $xf->format_key = 0;
            }

            $fmt = $this->format_map[$xf->format_key];
            $cellty = Defs::$cellty_from_fmtty[$fmt->type];

            $this->xf_index_to_xl_type_map[$xf->xf_index] = $cellty;
            # Now for some assertions etc
            if (!$this->formatting_info) {
                continue;
            }
            if ($xf->is_style) {
                continue;
            }

            if ($xf->parent_style_index < 0 || $xf->parent_style_index >= $num_xfs) {
                $xf->parent_style_index = 0;
            }

        }
    }

    public function namesEpilogue()
    {
        $num_names = count($this->name_obj_list);

        Helper::debug("+++++ names_epilogue +++++");
        Helper::debug("_all_sheets_map %s", join(', ', $this->all_sheets_map));
        Helper::debug("_extnsht_name_from_num %s", join(', ', $this->extnsht_name_from_num));
        Helper::debug("_sheet_num_from_name %s", join(', ', $this->sheet_num_from_name));
        for ($namex = 0; $namex < $num_names; $namex++) {
            $nobj = $this->name_obj_list[$namex];
            # Convert from excel_sheet_index to scope.
            # This is done here because in BIFF7 and earlier, the
            # BOUNDSHEET records (from which _all_sheets_map is derived)
            # come after the NAME records.

            $intl_sheet_index = -1;
            if ($this->biff_version >= 80) {
                $sheet_index = $nobj->excel_sheet_index;
                if ($sheet_index == 0) {
                    $intl_sheet_index = -1; # global
                } else if ($sheet_index >= 1 && $sheet_index <= strlen($this->all_sheets_map[$sheet_index - 1])) {
                    $intl_sheet_index = $this->all_sheets_map[$sheet_index - 1];
                    if ($intl_sheet_index == -1) { #maps to a macro or VBA sheet
                        $intl_sheet_index = -2;
                    }
                } else {
                    # huh?
                    $intl_sheet_index = -3; # invalid
                }
            } elseif ($this->biff_version >= 50 && $this->biff_version <= 70) {
                $sheet_index = $nobj->extn_sheet_num;
                if ($sheet_index == 0) {
                    $intl_sheet_index = -1; # global
                } else {
                    $sheet_name = $this->extnsht_name_from_num[$sheet_index];
                    $intl_sheet_index = isset($this->sheet_num_from_name[$sheet_name]) ? $this->sheet_num_from_name[$sheet_name] : -2;
                }
            }
            $nobj->scope = $intl_sheet_index;
            $this->name_obj_list[$namex] = $nobj;
        }

        for ($namex = 0; $namex < $num_names; $namex++) {
            $nobj = $this->name_obj_list[$namex];
            # Parse the formula ...
            if ($nobj->macro || $nobj->binary) {
                continue;
            }
            if ($nobj->evaluated) {
                continue;
            }
//            throw new XLSParserException("Formula evaluation not implemented");
            // @TODO: evaluate_name_formula($nobj, $namex);
        }

        Helper::debug("---------- name object dump ----------");
        for ($namex = 0; $namex < $num_names; $namex++) {
            $nobj = $this->name_obj_list[$namex];
            Helper::debug($nobj->dump("--- name[$namex] ---", ''));
        }
        Helper::debug("--------------------------------------");


        #
        # Build some dicts for access to the name objects
        #
        $name_and_scope_map = []; # (name.lower(), scope): Name_object
        $name_map = [];           # name.lower() : list of Name_objects (sorted in scope order)
        for ($namex = 0; $namex < $num_names; $namex++) {
            // @TODO
            $nobj = $this->name_obj_list[$namex];
            $name_lcase = mb_strtolower($nobj->name);
            $key = [$name_lcase, $nobj->scope];
//            $name_and_scope_map[$key] = $nobj;

            $this->name_obj_list[$namex] = $nobj;
//            throw new XLSParserException("TODO");
        }

    }

    public function handleSST($data)
    {
        Helper::debug("SST Processing");
        $nbt = strlen($data);
        $strlist = [$data];
        # <i
        list($uniquestrings) = array_values(unpack('i', substr($data, 4, 4)));
        Helper::debug("SST: unique strings: %d", $uniquestrings);
        while (true) {
            list($code, $nb, $data) = $this->readRecordPartsConditional(XL_CONTINUE);
            if ($code == Null) {
                break;
            }
            $nbt += $nb;
            Helper::debug("CONTINUE: adding %d bytes to SST -> %d", $nb, $nbt);
            $strlist[] = $data;
        }
        list($this->sharedstrings, $rt_runlist) = $this->unpackSSTTable($strlist, $uniquestrings);
        if ($this->formatting_info) {
            $this->rich_text_runlist_map = $rt_runlist;
        }
    }

    public function unpackSSTTable($datatab, $nstrings)
    {
        $datainx = 0;
        $ndatas = count($datatab);
        $data = $datatab[0];
        $datalen = strlen($data);
        $pos = 8;
        $strings = [];
        $richtext_runs = [];
        for ($tmp = 0; $tmp < $nstrings; $tmp++) {
            # <H
            list($nchars) = array_values(unpack('v', substr($data, $pos, 2)));
            $pos += 2;
            $options = ord($data[$pos]);
            $pos += 1;
            $rtcount = $phosz = 0;
            if ($options & 0x08) { # richtext
                # <H
                list($rtcount) = array_values(unpack('v', substr($data, $pos, 2)));
                $pos += 2;
            }

            if ($options & 0x04) { # phonetic
                # <i
                list($phosz) = array_values(unpack('i', substr($data, $pos, 4)));
                $pos += 4;
            }
            $accstrg = '';
            $charsgot = 0;
            while (true) {
                $charsneed = $nchars - $charsgot;
                if ($options & 0x01) {
                    # Uncompressed UTF-16
                    $charsavail = min(($datalen - $pos) >> 1, $charsneed);
                    $rawstrg = substr($data, $pos, 2 * $charsavail);
                    $accstrg .= mb_convert_encoding($rawstrg, 'utf-8', 'utf-16le');
                    $pos += 2 * $charsavail;
                } else {
                    # Note: this is COMPRESSED (not ASCII!) encoding!!!
                    $charsavail = min($datalen - $pos, $charsneed);
                    $rawstrg = substr($data, $pos, $charsavail);
                    $accstrg .= mb_convert_encoding($rawstrg, 'utf-8', 'latin1');
                    $pos += $charsavail;
                }
                $charsgot += $charsavail;
                if ($charsgot == $nchars) {
                    break;
                }
                $datainx += 1;
                $data = $datatab[$datainx];
                $datalen = strlen($data);
                $options = ord($data[0]);
                $pos = 1;
            }

            if ($rtcount) {
                $runs = [];
                for ($runindex = 0; $runindex < $rtcount; $runindex++) {
                    if ($pos == $datalen) {
                        $pos = 0;
                        $datainx ++;
                        $data = $datatab[$datainx];
                        $datalen = strlen($data);
                    }
                    # <HH
                    $runs[] = array_values(unpack('v2', substr($data, $pos, 4)));
                    $pos += 4;
                }
                $richtext_runs[count($strings)] = $runs;
            }

            $pos += $phosz; # size of the phonetic stuff to skip
            if ($pos >= $datalen) {
                # adjust to correct position in next record
                $pos = $pos - $datalen;
                $datainx ++;
                if ($datainx < $ndatas) {
                    $data = $datatab[$datainx];
                    $datalen = strlen($data);
                } else {
                    assert($tmp == $nstrings - 1);
                }
            }

            $strings[] = $accstrg;
        }
        return [$strings, $richtext_runs];
    }

    public function handleExternSheet($data)
    {
        $this->deriveEncoding(); # in case CODEPAGE record missing/out of order/wrong
        $this->extnsht_count ++; # for use as a 1-based index
        if ($this->biff_version >= 80) {
            # <H
            list($num_refs) = array_values(unpack('v', substr($data, 0, 2)));
            $bytes_reqd = $num_refs * 6 + 2;
            while (strlen($data) < $bytes_reqd) {
                helper::log("INFO: EXTERNSHEET needs %d bytes, have %d", $bytes_reqd, strlen($data));
                # length2 is unused in this code
                list($code2, /*$length2*/, $data2) = $this->readRecordParts();
                if ($code2 != XL_CONTINUE) {
                    throw new XLSParserException("Missing CONTINUE after EXTERNSHEET record");
                }
                $data .= $data2;
            }
            $pos = 2;
            for ($k = 0; $k < $num_refs; $k++) {
                # <HHH
                $info = array_values(unpack('v3', substr($data, $pos, 6)));
                list($ref_recordx, $ref_first_sheetx, $ref_last_sheetx) = $info;
                $this->externsheet_info[] = $info;
                $pos += 6;
                Helper::log("EXTERNSHEET(b8): k = %2d, record = %2d, first_sheet = %5d, last sheet = %5d", $k, $ref_recordx, $ref_first_sheetx, $ref_last_sheetx);
            }
        } else {
            # <BB
            list($nc, $ty) = array_values(unpack('C2', substr($data, 0, 2)));
            Helper::log("EXTERNSHEET(b7-):");
            Helper::debug("    " . Helper::as_hex($data));
            $msgs = [1 => "Encoded URL", 2 => "Current sheet!!", 3 => "Specific sheet in own doc't", 4 => "Nonspecific sheet in own doc't!!",];
            $msg = isset($msgs[$ty]) ? $msgs[$ty] : 'Not encoded';
            Helper::log("    %3d chars, type is %d (%s)" % $nc, $ty, $msg);
            if ($ty == 3) {
                $sheet_name = mb_convert_encoding(substr($data, 2, $nc), 'utf-8', 'utf-16le');
                $this->extnsht_name_from_num[$this->extnsht_count] = $sheet_name;
                Helper::debug($this->extnsht_name_from_num);
            }
            if ($ty <1 || $ty > 4) {
                $ty = 0;
            }
            $this->externsheet_type_b57[] = $ty;
        }
    }

    public function handleSupbook($data)
    {
        # aka EXTERNALBOOK in OOo docs
        $this->supbook_types[] = Null;
        Helper::debug("SUPBOOK:");
        Helper::debug(Helper::as_hex($data));

        # <H
        list($num_sheets) = array_values(unpack('v', substr($data, 0, 2)));
        Helper::debug("num_sheets = %d", $num_sheets);
        $sbn = $this->supbook_count;
        $this->supbook_types ++;
        if (substr($data, 2, 2) == "\x01\x04") {
            $this->supbook_types[-1] = SUPBOOK_INTERNAL;
            $this->supbook_locals_inx = $this->supbook_count - 1;
            Helper::debug("SUPBOOK[%d]: internal 3D refs; %d sheets", $sbn, $num_sheets);
            Helper::debug("    _all_sheets_map %s", join(', ', $this->all_sheets_map));
            return;
        }
        if (substr($data, 0, 4) == "\x01\x00\x01\x3A") {
            $this->supbook_types[-1] = SUPBOOK_ADDIN;
            $this->supbook_addins_inx = $this->supbook_count - 1;
            Helper::debug("SUPBOOK[%d]: add-in functions", $sbn);
            return;
        }
        list($url, $pos) = Helper::unpack_unicode_update_pos($data, 2, 2);
        if ($num_sheets == 0) {
            $this->supbook_types[-1] = SUPBOOK_DDEOLE;
            Helper::debug("SUPBOOK[%d]: DDE/OLE document = %s", $sbn, $url);
            return;
        }
        $this->supbook_types[-1] = SUPBOOK_EXTERNAL;
        Helper::debug("SUPBOOK[%d]: url = %s", $sbn, $url);
        $sheet_names = [];
        for ($x = 0; $x < $num_sheets; $x++) {
            # #### FIX ME ####
            # Should implement handling of CONTINUE record(s) ...
            list($shname, $pos) = Helper::unpack_unicode_update_pos($data, $pos, 2);
            $sheet_names[] = $shname;
            Helper::debug("  sheetx=%d namelen=%d name=%s (next pos=%d)", $x, strlen($shname), $shname, $pos);
        }
    }

    public function handleCountry($data) {
        # <HH
        $countries = array_values(unpack('v2', substr($data, 0, 4)));
        Helper::log("Countries: %s", join(', ', $countries));
        # Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        assert($this->countries === [0, 0] || $this->countries === $countries);
        $this->countries = $countries;
    }

    public function handleBoundSheet($data)
    {
        $bv = $this->biff_version;
        $this->deriveEncoding();
        Helper::debug("BOUNDSHEET: bv=%d data %s", $bv, Helper::as_hex($data));
        if ($bv == 45) { # BIFF4W
            #### Not documented in OOo docs ...
            # In fact, the *only* data is the name of the sheet.
            $sheet_name = Helper::unpack_string($data, 0, $this->encoding, 1);
            $visibility = 0;
            $sheet_type = XL_BOUNDSHEET_WORKSHEET; # guess, patch later
            if (strlen($this->sh_abs_posn) == 0) {
                $abs_posn = $this->sheetoffset + $this->base;
                # Note (a) this won't be used
                # (b) it's the position of the SHEETHDR record
                # (c) add 11 to get to the worksheet BOF record
            } else {
                $abs_posn = -1;
            }
        } else {

            # <iBB
            list ($offset, $visibility, $sheet_type) = array_values(unpack('Va/C2b', substr($data, 0, 6)));
            $abs_posn = $offset + $this->base; # because global BOF is always at posn 0 in the stream
            if ($bv < BIFF_FIRST_UNICODE) {
                $sheet_name = Helper::unpack_string($data, 6, $this->encoding, 1);
            } else {
                $sheet_name = Helper::unpack_unicode($data, 6, 1);
            }
        }

        Helper::debug("BOUNDSHEET: inx=%d vis=%s sheet_name=%s abs_posn=%d sheet_type=0x%02x",
            $this->all_sheets_count, $visibility, $sheet_name, $abs_posn, $sheet_type);

        $this->all_sheets_count ++;
        if ($sheet_type != XL_BOUNDSHEET_WORKSHEET) {
            $this->all_sheets_map[] = -1;
            $descrs = [
                1 => 'Macro sheet',
                2 => 'Chart',
                6 => 'Visual Basic module',
            ];

            Helper::debug("NOTE *** Ignoring non-worksheet data named %s (type 0x%02x = %s)", $sheet_name, $sheet_type, isset($descrs[$sheet_type]) ? $descrs[$sheet_type] : 'UNKNOWN');
        } else {
            $snum = count($this->sheet_names);
            $this->all_sheets_map[] = $snum;

            $this->sheet_names[] = $sheet_name;
            $this->sh_abs_posn[] = $abs_posn;
            $this->sheet_visibility[] = $visibility;
            $this->sheet_num_from_name[$sheet_name] = $snum;
        }
    }

    public function handleDateMode($data)
    {
        # <H
        list($datemode) = array_values(unpack('v', substr($data, 0, 2)));
        Helper::debug("DATEMODE: datemode %s", $datemode);
        assert($datemode === 0 || $datemode === 1);
        $this->datemode = $datemode;
    }

    public function handleCodePage($data)
    {
        # <H
        list($codepage) = array_values(unpack('v', substr($data, 0, 2)));
        $this->codepage = $codepage;
        $this->deriveEncoding();
    }

    public function deriveEncoding()
    {
        $encoding_from_codepage = [
            1200  => 'utf_16_le',
            10000 => 'mac_roman',
            10006 => 'mac_greek', # guess
            10007 => 'mac_cyrillic', # guess
            10029 => 'mac_latin2', # guess
            10079 => 'mac_iceland', # guess
            10081 => 'mac_turkish', # guess
            32768 => 'mac_roman',
            32769 => 'cp1252',
            1257 => 'iso-8859-13',
            1252 => 'windows-1252',
            1258 => 'windows-1252',
            1253 => 'iso-8859-7',
        ];
        if ($this->encoding_override) {
            $this->encoding = $this->encoding_override;
        } elseif ($this->codepage === null) {
            if ($this->biff_version < 80) {
                Helper::log("*** No CODEPAGE record, no encoding_override: will use 'ascii'");
                $this->encoding = 'ascii';
            } else {
                $this->codepage = 1200; #utf16le
                Helper::log("*** No CODEPAGE record; assuming 1200 (utf_16_le)");
            }
        } else {
            $codepage = $this->codepage;
            if (isset($encoding_from_codepage[$codepage])) {
                $encoding = $encoding_from_codepage[$codepage];
            } elseif ($codepage >= 300 && $codepage <= 1999) {
                $encoding = 'cp' . $codepage;
            } else {
                $encoding = "unknown_codepage_" . $codepage;
            }
            $this->encoding = $encoding;
            if ($this->encoding != $encoding) {
                Helper::log("CODEPAGE: codepage %s -> encoding %s", $codepage, $encoding);
            } else {
                Helper::debug("CODEPAGE: codepage %s -> encoding %s", $codepage, $encoding);
            }
        }

        if ($this->raw_user_name) {
            $strg = Helper::unpack_string($this->user_name, 0, $this->encoding, 1);
            $strg = rtrim($strg);
            $this->user_name = $strg;
            $this->raw_user_name = false;
        }

        return $this->encoding;

    }

    public function readRecordParts()
    {
        $pos = $this->position;
        $data = $this->mem;
        # <HH
        list($code, $length) = array_values(unpack('v2', substr($data, $pos, 4)));
        $pos += 4;
        $data = substr($data, $pos, $length);
        $this->position = $pos + $length;

        return [$code, $length, $data];
    }

    public function readRecordPartsConditional($reqd_record)
    {
        $pos = $this->position;
        $data = $this->mem;
        # <HH
        list($code, $length) = array_values(unpack('v2', substr($data, $pos, 4)));
        if ($code != $reqd_record) {
            return [Null, 0, ''];
        }
        $pos += 4;
        $data = substr($data, $pos, $length);
        $this->position = $pos + $length;
        return [$code, $length, $data];
    }

    public function loadBiff28()
    {
        $this->base = 0;
        $cd = new CompDoc($this->mem);

        if (USE_FANCY_CD) {
            foreach (['Workbook', 'Book'] as $qname) {
                list($this->mem, $this->base, $this->stream_len) = $cd->locate_named_stream($qname);

                if ($this->mem) break;
            }
            if (!$this->mem) {
                throw new XLSParserException("Can't find workbook in OLE2 compound document");
            }
        } else {
            foreach (['Workbook', 'Book'] as $qname) {
                $this->mem = $cd->get_named_stream($qname);
                if ($this->mem) break;
            }
            if (!$this->mem) {
                throw new XLSParserException("Can't find workbook in OLE2 compound document");
            }
            $this->stream_len = strlen($this->mem);
        }
        $this->position = $this->base;
    }

    public function readBOF($rqd_stream)
    {
        Helper::debug("reqd: 0x%04x", $rqd_stream);
        $savpos = $this->position;
        $opcode = $this->read2Bytes();
        if ($opcode == MY_EOF) {
            throw new XLSParserException("Unsupported format, or corrupt file: Expected BOF record; met end of file");
        }
        if (!in_array($opcode, Defs::$bofcodes)) {
            throw new XLSParserException("Expected BOF record; found " . substr($this->mem, $savpos, 8));
        }
        $length = $this->read2Bytes();
        if ($length == MY_EOF) {
            throw new XLSParserException("Expected BOF record[1]; met end of file");
        }
        if ($length < 4 || $length > 20) {
            throw new XLSParserException("Invalid length ($length) for BOF record type 0x" . dechex($opcode));
        }
        $padding = str_repeat("\x00", max(0, Defs::$boflen[$opcode] - $length));
        $data = $this->read($this->position, $length);
        Helper::debug("getbof(): data=%s", Helper::as_hex($data));
        if (strlen($data) < $length) {
            throw new XLSParserException("Incomplete BOF record[2]; met end of file");
        }

        $data .= $padding;
        $version1 = $opcode >>8;
        # <HH
        list($version2, $streamtype) = array_values(unpack('v2', substr($data, 0, 4)));
        Helper::debug("getbof(): op=0x%04x version2=0x%04x streamtype=0x%04x", $opcode, $version2, $streamtype);
        $bof_offset = $this->position - 4 - $length;
        Helper::debug("getbof(): BOF found at offset %d; savpos=%d", $bof_offset, $savpos);
        $version = $build = $year = 0;
        if ($version1 == 0x08) {
            # <HH
            list($build, $year) = array_values(unpack('v2', substr($data, 4, 4)));
            if ($version2 == 0x0600) {
                $version = 80;
            } elseif ($version2 == 0x0500) {
                if ($year < 1994 || in_array($build, [2412, 3218, 3321])) {
                    $version = 50;
                } else {
                    $version = 70;
                }
            } else {
                # dodgy one, created by a 3rd party tool
                $arr = [
                    0x0000 => 21,
                    0x0007 => 21,
                    0x0200 => 21,
                    0x0300 => 30,
                    0x0400 => 40,
                ];
                $version = isset($arr[$version2]) ? $arr[$version2] : 0;
            }
        } elseif (in_array($version1, [0x04, 0x02, 0x00])) {
            $arr = [0x04 => 40, 0x02 => 30, 0x00 => 21];
            $version = $arr[$version1];
        }

        if ($version == 40 && $streamtype == XL_WORKBOOK_GLOBALS_4W) {
            $version = 45; # i.e. 4W
        }

        Helper::debug("BOF: op=0x%04x vers=0x%04x stream=0x%04x buildid=%d buildyr=%d -> BIFF%d", $opcode, $version2, $streamtype, $build, $year, $version);

        $got_globals = $streamtype == XL_WORKBOOK_GLOBALS || ($version == 45 && $streamtype == XL_WORKBOOK_GLOBALS_4W);

        if (($rqd_stream == XL_WORKBOOK_GLOBALS && $got_globals) || $streamtype == $rqd_stream) {
            return $version;
        }
        if ($version < 50 && $streamtype == XL_WORKSHEET) {
            return $version;
        }
        if ($version >= 50 && $streamtype == 0x0100) {
            throw new XLSParserException("Workspace file - no spreadsheet data");
        }
        throw new XLSParserException(sprintf('BOF not workbook/worksheet: op=0x%s vers=0x%s strm=0x%s build = %d year = %d -> BIFF = %d', dechex($opcode), dechex($version2), dechex($streamtype), $build, $year, $version));
    }

    public function read2Bytes()
    {
        $pos = $this->position;
        $buff_two = substr($this->mem, $pos, 2);
        $lenbuff = strlen($buff_two);
        $this->position += $lenbuff;
        if ($lenbuff < 2) {
            return MY_EOF;
        }
        # <H
        list($ret) = array_values(unpack('v', $buff_two));
        return $ret;
    }

    public function read($pos, $length)
    {
        $data = substr($this->mem, $pos, $length);
        $this->position = $pos + strlen($data);
        return $data;
    }

    public function dumpAsHex($str)
    {
        return join(' ', array_map(function($el){
            return '\x'.str_pad(dechex(ord($el)), 2, '0', STR_PAD_LEFT);
        }, str_split($str)));

    }

    function sheets()
    {
        $ret = [];
        foreach ($this->sheet_list as $sheet) {
            $ret[$sheet->name] = $sheet;
        }
        return $ret;
    }

    public function getIterator()
    {
        return new ArrayIterator($this->sheet_list);
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
        if ($offset === (int)$offset) {
            return isset($this->sheet_list[$offset]);
        } else {
            return isset($this->sheet_names[$offset]);
        }
    }

    /**
     * Offset to retrieve
     * @link http://php.net/manual/en/arrayaccess.offsetget.php
     * @param mixed $offset The offset to retrieve.
     * @return mixed Can return all value types.
     */
    public function offsetGet($offset)
    {
        if ($offset === (int)$offset) {
            return $this->getSheetByIndex($offset);
        } else {
            return $this->getSheetByName($offset);
        }
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
