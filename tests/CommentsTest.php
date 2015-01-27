<?php

use laacz\XLSParser\Book;

class CommentsTest extends PHPUnit_Framework_TestCase {

    /**
     * @var Book
     */
    public $book;

    function setUp()
    {
        $this->book = new Book(file_get_contents(__DIR__ . '/xls/comments.xls'));

    }

    function testComments()
    {
        $sheet = $this->book->getSheetByName('Blätt1');
        $note = $sheet->cell_note_map[1][1];
        $this->assertEquals("laacz:\nComment (ASCII)", $note->text);
        $note = $sheet->cell_note_map[4][2];
        $this->assertEquals("Glāžšķūnis:\nKomentārs (glāžšķūņi)", $note->text);

        $sheet = $this->book->getSheetByName('Formate');
        $note = $sheet->cell_note_map[1][0];
        $this->assertEquals("Komentārs\nadasd", $note->text);
    }

    function testCommentsFormatting()
    {
        $sheet = $this->book->getSheetByName('Formate');
        $note = $sheet->cell_note_map[1][0];
        $this->assertEquals([[0, 11], [11, 31], [12, 32], [13, 10]], $note->rich_text_runlist);
    }

}
