<?php

namespace laacz\XLSParser;

class Defs
{

    /**
     * @var array
     */
    static $supported_versions = [80, 70, 50, 45, 40, 30, 21, 20];
    /**
     * @var array
     */
    static $biff_text_from_num = [
        '0'  => "(not BIFF)",
        '20' => "2.0",
        '21' => "2.1",
        '30' => "3",
        '40' => "4S",
        '45' => "4W",
        '50' => "5",
        '70' => "7",
        '80' => "8",
        '85' => "8X",
    ];

    /**
     * @var array
     */
    static $std_format_strings = [
        # "std" == "standard for US English locale"
        # #### TODO ... a lot of work to tailor these to the user's locale.
        # See e.g. gnumeric-1.x.y/src/formats.c
        0x00 => "General",
        0x01 => "0",
        0x02 => "0.00",
        0x03 => "#,##0",
        0x04 => "#,##0.00",
        0x05 => "$#,##0_);($#,##0)",
        0x06 => "$#,##0_);[Red]($#,##0)",
        0x07 => "$#,##0.00_);($#,##0.00)",
        0x08 => "$#,##0.00_);[Red]($#,##0.00)",
        0x09 => "0%",
        0x0a => "0.00%",
        0x0b => "0.00E+00",
        0x0c => "# ?/?",
        0x0d => "# ??/??",
        0x0e => "m/d/yy",
        0x0f => "d-mmm-yy",
        0x10 => "d-mmm",
        0x11 => "mmm-yy",
        0x12 => "h:mm AM/PM",
        0x13 => "h:mm:ss AM/PM",
        0x14 => "h:mm",
        0x15 => "h:mm:ss",
        0x16 => "m/d/yy h:mm",
        0x25 => "#,##0_);(#,##0)",
        0x26 => "#,##0_);[Red](#,##0)",
        0x27 => "#,##0.00_);(#,##0.00)",
        0x28 => "#,##0.00_);[Red](#,##0.00)",
        0x29 => "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
        0x2a => "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
        0x2b => "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
        0x2c => "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        0x2d => "mm:ss",
        0x2e => "[h]:mm:ss",
        0x2f => "mm:ss.0",
        0x30 => "##0.0E+0",
        0x31 => "@",
    ];

    /**
     * @var array
     */
    static $fmt_code_ranges = [ # both-inclusive ranges of "standard" format codes
        # Source: the openoffice.org doc't
        # and the OOXML spec Part 4, section 3.8.30
        [ 0,  0, FGE],
        [ 1, 13, FNU],
        [14, 22, FDT],
        [27, 36, FDT], # CJK date formats
        [37, 44, FNU],
        [45, 47, FDT],
        [48, 48, FNU],
        [49, 49, FTX],
        # Gnumeric assumes (or assumed) that built-in formats finish at 49, not at 163
        [50, 58, FDT], # CJK date formats
        [59, 62, FNU], # Thai number (currency?) formats
        [67, 70, FNU], # Thai number (currency?) formats
        [71, 81, FDT], # Thai date formats
    ];
    /**
     * @var array
     */
    static $std_format_code_types = [];

    /**
     * @var array
     */
    static $excel_default_palette_b5 = [
        [  0,   0,   0], [255, 255, 255], [255,   0,   0], [  0, 255,   0],
        [  0,   0, 255], [255, 255,   0], [255,   0, 255], [  0, 255, 255],
        [128,   0,   0], [  0, 128,   0], [  0,   0, 128], [128, 128,   0],
        [128,   0, 128], [  0, 128, 128], [192, 192, 192], [128, 128, 128],
        [153, 153, 255], [153,  51, 102], [255, 255, 204], [204, 255, 255],
        [102,   0, 102], [255, 128, 128], [  0, 102, 204], [204, 204, 255],
        [  0,   0, 128], [255,   0, 255], [255, 255,   0], [  0, 255, 255],
        [128,   0, 128], [128,   0,   0], [  0, 128, 128], [  0,   0, 255],
        [  0, 204, 255], [204, 255, 255], [204, 255, 204], [255, 255, 153],
        [153, 204, 255], [255, 153, 204], [204, 153, 255], [227, 227, 227],
        [ 51, 102, 255], [ 51, 204, 204], [153, 204,   0], [255, 204,   0],
        [255, 153,   0], [255, 102,   0], [102, 102, 153], [150, 150, 150],
        [  0,  51, 102], [ 51, 153, 102], [  0,  51,   0], [ 51,  51,   0],
        [153,  51,   0], [153,  51, 102], [ 51,  51, 153], [ 51,  51,  51],
    ];

    /**
     * @var array
     */
    static $excel_default_palette_b2 = [
        [  0,   0,   0], [255, 255, 255], [255,   0,   0], [  0, 255,   0],
        [  0,   0, 255], [255, 255,   0], [255,   0, 255], [  0, 255, 255],
        [128,   0,   0], [  0, 128,   0], [  0,   0, 128], [128, 128,   0],
        [128,   0, 128], [  0, 128, 128], [192, 192, 192], [128, 128, 128]
    ];

    # Following table borrowed from Gnumeric 1.4 source.
    # Checked against OOo docs and MS docs.
    /**
     * @var array
     */
    static $excel_default_palette_b8 = [ # [red, green, blue]
        [  0,  0,  0], [255,255,255], [255,  0,  0], [  0,255,  0], # 0
        [  0,  0,255], [255,255,  0], [255,  0,255], [  0,255,255], # 4
        [128,  0,  0], [  0,128,  0], [  0,  0,128], [128,128,  0], # 8
        [128,  0,128], [  0,128,128], [192,192,192], [128,128,128], # 12
        [153,153,255], [153, 51,102], [255,255,204], [204,255,255], # 16
        [102,  0,102], [255,128,128], [  0,102,204], [204,204,255], # 20
        [  0,  0,128], [255,  0,255], [255,255,  0], [  0,255,255], # 24
        [128,  0,128], [128,  0,  0], [  0,128,128], [  0,  0,255], # 28
        [  0,204,255], [204,255,255], [204,255,204], [255,255,153], # 32
        [153,204,255], [255,153,204], [204,153,255], [255,204,153], # 36
        [ 51,102,255], [ 51,204,204], [153,204,  0], [255,204,  0], # 40
        [255,153,  0], [255,102,  0], [102,102,153], [150,150,150], # 44
        [  0, 51,102], [ 51,153,102], [  0, 51,  0], [ 51, 51,  0], # 48
        [153, 51,  0], [153, 51,102], [ 51, 51,153], [ 51, 51, 51], # 52
    ];

    # 00H = Normal
    # 01H = RowLevel_lv (see next field)
    # 02H = ColLevel_lv (see next field)
    # 03H = Comma
    # 04H = Currency
    # 05H = Percent
    # 06H = Comma [0] (BIFF4-BIFF8)
    # 07H = Currency [0] (BIFF4-BIFF8)
    # 08H = Hyperlink (BIFF8)
    # 09H = Followed Hyperlink (BIFF8)
    /**
     * @var array
     */
    static $built_in_style_names = [
        "Normal",
        "RowLevel_",
        "ColLevel_",
        "Comma",
        "Currency",
        "Percent",
        "Comma [0]",
        "Currency [0]",
        "Hyperlink",
        "Followed Hyperlink",
    ];

    /**
     * @var array
     */
    static $default_palette = [
        80 => 'excel_default_palette_b8',
        70 => 'excel_default_palette_b5',
        50 => 'excel_default_palette_b5',
        45 => 'excel_default_palette_b2',
        40 => 'excel_default_palette_b2',
        30 => 'excel_default_palette_b2',
        21 => 'excel_default_palette_b2',
        20 => 'excel_default_palette_b2',
    ];

    /**
     * This dictionary can be used to produce a text version of the internal codes
     * that Excel uses for error cells. Here are its contents:
     *
     * 0x00: '#NULL!',  # Intersection of two cell ranges is empty
     * 0x07: '#DIV/0!', # Division by zero
     * 0x0F: '#VALUE!', # Wrong type of operand
     * 0x17: '#REF!',   # Illegal or deleted cell reference
     * 0x1D: '#NAME?',  # Wrong function or range name
     * 0x24: '#NUM!',   # Value range overflow
     * 0x2A: '#N/A',    # Argument or function not available
     *
     * @var array
     */
    static $error_text_from_code = [
        0x00 => '#NULL!',  # Intersection of two cell ranges is empty
        0x07 => '#DIV/0!', # Division by zero
        0x0F => '#VALUE!', # Wrong type of operand
        0x17 => '#REF!',   # Illegal or deleted cell reference
        0x1D => '#NAME?',  # Wrong function or range name
        0x24 => '#NUM!',   # Value range overflow
        0x2A => '#N/A',    # Argument or function not available
    ];

    /**
     * @var array
     */
    static $cellty_from_fmtty = [
        FNU => XL_CELL_NUMBER,
        FUN => XL_CELL_NUMBER,
        FGE => XL_CELL_NUMBER,
        FDT => XL_CELL_DATE,
        FTX => XL_CELL_NUMBER, # Yes, a number can be formatted as text.
    ];

    /**
     * @var array
     */
    static $WINDOW2_options = [
        # Attribute names and initial values to use in case
        # a WINDOW2 record is not written.
        "show_formulas" => 0,
        "show_grid_lines" => 1,
        "show_sheet_headers" => 1,
        "panes_are_frozen" => 0,
        "show_zero_values" => 1,
        "automatic_grid_line_colour" => 1,
        "columns_from_right_to_left" => 0,
        "show_outline_symbols" => 1,
        "remove_splits_if_pane_freeze_is_removed" => 0,
        # Multiple sheets can be selected, but only one can be active
        # (hold down Ctrl and click multiple tabs in the file in OOo)
        "sheet_selected" => 0,
        # "sheet_visible" should really be called "sheet_active"
        # and is 1 when this sheet is the sheet displayed when the file
        # is open. More than likely only one sheet should ever be set as
        # visible.
        # This would correspond to the Book's sheet_active attribute, but
        # that doesn't exist as WINDOW1 records aren't currently processed.
        # The real thing is the visibility attribute from the BOUNDSHEET record.
        "sheet_visible" => 0,
        "show_in_page_break_preview" => 0,
    ];

    static $boflen = [0x0809 => 8, 0x0409 => 6, 0x0209 => 6, 0x0009 => 4];
    static $bofcodes = [0x0809, 0x0409, 0x0209, 0x0009];

    static $code_from_builtin_name = [
        "Consolidate_Area" =>"\x00",
        "Auto_Open" =>       "\x01",
        "Auto_Close" =>      "\x02",
        "Extract" =>         "\x03",
        "Database" =>        "\x04",
        "Criteria" =>        "\x05",
        "Print_Area" =>      "\x06",
        "Print_Titles" =>    "\x07",
        "Recorder" =>        "\x08",
        "Data_Form" =>       "\x09",
        "Auto_Activate" =>   "\x0A",
        "Auto_Deactivate" => "\x0B",
        "Sheet_Title" =>     "\x0C",
        "_FilterDatabase" => "\x0D",
    ];


    static $builtin_name_from_code = [
        "\x00" => "Consolidate_Area",
        "\x01" => "Auto_Open",
        "\x02" => "Auto_Close",
        "\x03" => "Extract",
        "\x04" => "Database",
        "\x05" => "Criteria",
        "\x06" => "Print_Area",
        "\x07" => "Print_Titles",
        "\x08" => "Recorder",
        "\x09" => "Data_Form",
        "\x0A" => "Auto_Activate",
        "\x0B" => "Auto_Deactivate",
        "\x0C" => "Sheet_Title",
        "\x0D" => "_FilterDatabase",
    ];


}