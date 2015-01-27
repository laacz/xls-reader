<?php


use laacz\XLSParser\Book;
use laacz\XLSParser\Sheet;

class FormulasTest extends \PHPUnit_Framework_TestCase {

    /**
     * @var Book
     */
    var $book;

    /**
     * @var Sheet
     */
    var $sheet;

    public function setUp()
    {
    }

    public function testFormulas()
    {
        $this->book = new Book(file_get_contents(__DIR__ . '/xls/formula_test_sjmachin.xls'));
        $this->sheet = $this->book[0];

        $this->assertEquals("МОСКВА Москва", (string)$this->sheet[1][1], 'B2');
        $this->assertEquals(0.14285714285714285, $this->sheet[2][1]->value, 'B3', .0000000001);
        $this->assertEquals("ABCDEF", (string)$this->sheet[3][1], 'B4');
        $this->assertEquals("", (string)$this->sheet[4][1], 'B5');
        $this->assertEquals('1', (string)$this->sheet[5][1], 'B6');
        $this->assertEquals('7', (string)$this->sheet[6][1], 'B7');
        $this->assertTrue((string)$this->sheet[1][1] == (string)$this->sheet[7][1], 'Cells B2 and B8 should be equal');
    }

    public function testNameFormulas()
    {
        $this->book = new Book(file_get_contents(__DIR__. '/xls/formula_test_names.xls'));
        $this->sheet = $this->book[0];

        $this->assertEquals('-7.0', $this->sheet[1][1]->value);
        $this->assertEquals('4.0', $this->sheet[2][1]->value);
        $this->assertEquals('6.0', $this->sheet[3][1]->value);
        $this->assertEquals('3.0', $this->sheet[4][1]->value);
        $this->assertEquals("b", $this->sheet[5][1]->value);
        $this->assertEquals("C", $this->sheet[6][1]->value);

    }
}
