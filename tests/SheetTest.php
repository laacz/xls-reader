<?php

use laacz\XLSParser\Book;
use laacz\XLSParser\Cell;
use laacz\XLSParser\Sheet;

define('SHEETINDEX', 0);
define('NROWS', 21);
define('NCOLS', 13);

define('ROW_ERR', NROWS + 10);
define('COL_ERR', NCOLS + 10);


class SheetTest extends PHPUnit_Framework_TestCase {

    /**
     * @var Book
     */
    private $book;

    /**
     * @var Sheet
     */
    private $sheet;

    private $sheetnames = ['PROFILEDEF', 'AXISDEF', 'TRAVERSALCHAINAGE', 'AXISDATUMLEVELS', 'PROFILELEVELS'];

    private $cells = [
        ['PROFIL',  'a',   'b',   'c',   'd',   'e',   'f',   'g',   'h',   'i',   'j',   'k',   'l'],
        ['P8.2',    '100', '101', '102', '103', '104', '105', '106', '107', '108', '109', '110', '111'],
        ['P8.3',    '112', '113', '114', '115', '116', '117', '118', '119', '120', '121', '122', '123'],
        ['P8.4',    '124', '125', '126', '127', '128', '129', '130', '131', '132', '133', '134', '135'],
        ['P8.5',    '136', '137', '138', '139', '140', '141', '142', '143', '144', '145', '146', '147'],
        ['P8.6',    '148', '149', '150', '151', '152', '153', '154', '155', '156', '157', '158', '159'],
        ['P8.7',    '160', '161', '162', '163', '164', '165', '166', '167', '168', '169', '170', '171'],
        ['P9',      '172', '173', '174', '175', '176', '177', '178', '179', '180', '181', '182', '183'],
        ['P9.1',    '184', '185', '186', '187', '188', '189', '190', '191', '192', '193', '194', '195'],
        ['P9.2',    '196', '197', '198', '199', '200', '201', '202', '203', '204', '205', '206', '207'],
        ['P9.3',    '208', '209', '210', '211', '212', '213', '214', '215', '216', '217', '218', '219'],
        ['P9.4',    '220', '221', '222', '223', '224', '225', '226', '227', '228', '229', '230', '231'],
        ['P9.5',    '232', '233', '234', '235', '236', '237', '238', '239', '240', '241', '242', '243'],
        ['P9.6',    '244', '245', '246', '247', '248', '249', '250', '251', '252', '253', '254', '255'],
        ['Q0',      '256', '257', '258', '259', '260', '261', '262', '263', '264', '265', '266', '267'],
        [''],
        [''],
        ['', '', '2014-12-01', ],
        ['', '', '12:03'],
        ['', '', '2015-12-01 13:37:00'],
        ['', '', '1900-01-01'],
    ];

    public function setUp()
    {
        $this->book = new Book(file_get_contents(__DIR__ . '/xls/profiles.xls'));
        $this->sheet = $this->book->getSheetByIndex(SHEETINDEX);
    }

    public function testSheetsAccessByIndex()
    {
        $this->assertEquals('PROFILEDEF', $this->book['PROFILEDEF']->name);
        $this->assertEquals('PROFILEDEF', $this->book[0]->name);
        $this->assertEquals('AXISDATUMLEVELS', $this->book[3]->name);
    }

    public function testSheetsIterator()
    {
        $sheetnames = [];
        foreach ($this->book as $sheet) {
            $sheetnames[] = $sheet->name;
        }
        $this->assertEquals($this->sheetnames, $sheetnames);
    }

    public function testNRows()
    {
        $this->assertEquals(NROWS, $this->sheet->nrows);
    }

    public function testNCols()
    {
        $this->assertEquals(NCOLS, $this->sheet->ncols);
    }

    public function testCell()
    {
        $empty = new Cell(XL_CELL_EMPTY, '');

        $this->assertNotEquals($empty->ctype, $this->sheet->cell(0, 0)->ctype);
        $this->assertNotEquals($empty->value, $this->sheet->cell(0, 0)->value);

        $this->assertEquals($empty->ctype, $this->sheet->cell(NROWS - 1, NCOLS - 1)->ctype);
        $this->assertEquals($empty->value, $this->sheet->cell(NROWS - 1, NCOLS - 1)->value);
    }

    public function testCellType()
    {
        $this->assertEquals(XL_CELL_TEXT, $this->sheet->cell_type(0, 0));
        $this->assertEquals(XL_CELL_NUMBER, $this->sheet->cell_type(1, 1));
        $this->assertEquals(XL_CELL_NUMBER, $this->sheet->cell_type(2, 2));
        $this->assertEquals(XL_CELL_DATE, $this->sheet->cell_type(17, 2));
    }

    public function testCellValue()
    {
        $this->assertEquals('PROFIL', $this->sheet->cell_value(0, 0));
        $this->assertEquals(100, $this->sheet->cell_value(1, 1));
        $this->assertEquals(113, $this->sheet->cell_value(2, 2));
        $this->assertEquals("2014-12-01 00:00:00", $this->sheet->cell_value(17, 2));
        $this->assertEquals('12:03:00', $this->sheet->cell_value(18, 2));
        $this->assertEquals('2015-12-01 13:37:01', $this->sheet->cell_value(19, 2));
        $this->assertEquals('1900-01-01 00:00:00', $this->sheet->cell_value(20, 2));
        $sheet = $this->book->getSheetByName('PROFILELEVELS');
        $this->assertEquals(.025, $sheet->cell_value(1, 1), null, .001);
        $this->assertEquals(265.212, $sheet->cell_value(1, 2), null, .001);

        $this->assertEquals('PROFIL', $this->sheet[0][0]->value);
        $this->assertEquals(100, $this->sheet[1][1]->value);
        $this->assertEquals(113, $this->sheet[2][2]->value);
        $this->assertEquals("2014-12-01 00:00:00", $this->sheet[17][2]->value);
        $this->assertEquals('12:03:00', $this->sheet[18][2]->value);
        $this->assertEquals('2015-12-01 13:37:01', $this->sheet[19][2]->value);
        $this->assertEquals('1900-01-01 00:00:00', $this->sheet[20][2]->value);
        $sheet = $this->book->getSheetByName('PROFILELEVELS');
        $this->assertEquals(.025, $sheet[1][1]->value, null, .001);
        $this->assertEquals(265.212, $sheet[1][2]->value, null, .001);

    }

    public function testRowsIterator()
    {
        foreach ($this->sheet as $rowx=>$row) {
            $this->assertEquals($this->cells[$rowx][0], $this->sheet[$rowx][0]->value);
        }
    }


    public function testCol()
    {
    }
}
