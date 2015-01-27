<?php

use laacz\XLSParser\Book;
use laacz\XLSParser\Sheet;

class CellTest extends PHPUnit_Framework_TestCase {

    /**
     * @var Book
     */
    private $book;

    /**
     * @var Sheet
     */
    private $sheet;

    public function setUp()
    {
        $this->book = new Book(file_get_contents(__DIR__ . '/xls/profiles.xls'));
        $this->sheet = $this->book->getSheetByName('PROFILEDEF');
    }

    public function testStringCell()
    {
        $cell = $this->sheet->cell(0, 0);

        $this->assertEquals(XL_CELL_TEXT, $cell->ctype);
        $this->assertEquals('PROFIL', $cell->value);
        $this->assertTrue($cell->xf_index > 0);
    }

    public function testNumberCell()
    {
        $cell = $this->sheet->cell(1, 1);

        $this->assertEquals(XL_CELL_NUMBER, $cell->ctype);
        $this->assertEquals(100, $cell->value);
        $this->assertTrue($cell->xf_index > 0);
    }

    public function testCalculatedCell()
    {
        $sheet2 = $this->book->getSheetByName('PROFILELEVELS');
        $cell = $sheet2->cell(1, 3);
        $this->assertEquals(XL_CELL_NUMBER ,$cell->ctype);
        $this->assertEquals(265.131, $cell->value, '', .001);
        $this->assertTrue($cell->xf_index > 0);
    }

    public function testMergedCells()
    {
        $book = new Book(file_get_contents(__DIR__ . '/xls/merged.xls'));
        $sheet = $book->getSheetByName('Sheet1');
        $this->assertEquals([[0, 1, 1, 4], [1, 2, 0, 2], [0, 2, 4, 5]], $sheet->merged_cells);
        $this->assertEquals(XL_CELL_BLANK, $sheet[0][0]->ctype);
        $this->assertEquals(XL_CELL_BLANK, $sheet[0][2]->ctype);
        $this->assertEquals(XL_CELL_BLANK, $sheet[1][2]->ctype);
        $this->assertEquals(XL_CELL_BLANK, $sheet[1][4]->ctype);

        $this->assertEquals(XL_CELL_TEXT, $sheet[0][1]->ctype);
        $this->assertEquals(XL_CELL_TEXT, $sheet[1][0]->ctype);
        $this->assertEquals(XL_CELL_TEXT, $sheet[1][3]->ctype);
        $this->assertEquals(XL_CELL_NUMBER, $sheet[0][4]->ctype);

        $this->assertEquals('MERGED A1:D1', $sheet[0][1]->value);
        $this->assertEquals('MERGED A2:B2', $sheet[1][0]->value);
        $this->assertEquals('D2', $sheet[1][3]->value);
        $this->assertEquals(1.0, $sheet[0][4]->value, '', .001);
    }
}
