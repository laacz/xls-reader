<?php

use laacz\XLSParser\Book;
use laacz\XLSParser\Sheet;

class FormatTest extends \PHPUnit_Framework_TestCase {

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
        $this->book = new Book(file_get_contents(__DIR__ . '/xls/Formate.xls'));
        $this->sheet = $this->book->getSheetByName('Blätt1');
    }

    public function testTextCells()
    {
        foreach (['Huber', 'Äcker', 'Öcker'] as $row=>$name) {
            $cell = $this->sheet->cells[$row][0];

            $this->assertEquals(XL_CELL_TEXT, $cell->ctype);
            $this->assertEquals($name, $cell->value);
            $this->assertTrue($cell->xf_index > 0);
        }
    }

    public function testDateCells()
    {
        foreach ([
                     [0, '1907-07-03 00:00:00'], [1, '2005-02-23 00:00:00'], [2, '1988-05-03 00:00:00'],
                     [3, '06:34:00'], [4, '12:56:00'], [5, '17:47:13']
                 ] as $row) {
            list($row, $date) = $row;
            $cell = $this->sheet[$row][1];

            $this->assertEquals(XL_CELL_DATE, $cell->ctype);
            $this->assertEquals($date, $cell->value);
            $this->assertTrue($cell->xf_index > 0);
        }
    }

    public function testPercentCells()
    {
        foreach ([[6, .974], [7, .124],] as $row) {
            list($row, $number) = $row;

            $cell = $this->sheet[$row][1];

            $this->assertEquals(XL_CELL_NUMBER, $cell->ctype);
            $this->assertEquals($number, $cell->value);
            $this->assertTrue($cell->xf_index > 0);
        }

    }

    public function testCurrencyCells()
    {
        foreach ([[8, 1000.30], [9, 1.20],] as $row) {
            list($row, $number) = $row;

            $cell = $this->sheet[$row][1];

            $this->assertEquals(XL_CELL_NUMBER, $cell->ctype);
            $this->assertEquals($number, $cell->value);
            $this->assertTrue($cell->xf_index > 0);
        }

    }

    public function testGetFromMergedCell()
    {
        $cell = $this->book['ÖÄÜ'][2][2];

        $this->assertEquals(XL_CELL_TEXT, $cell->ctype);
        $this->assertEquals('MERGED CELLS', $cell->value);
        $this->assertTrue($cell->xf_index > 0);
    }

    public function testIgnoreDiagram()
    {
        $cell = $this->book['Blätt3'][0][0];

        $this->assertEquals(XL_CELL_NUMBER, $cell->ctype);
        $this->assertEquals(100, $cell->value);
        $this->assertTrue($cell->xf_index > 0);
    }

}
