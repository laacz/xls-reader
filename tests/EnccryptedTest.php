<?php

use laacz\XLSParser\Book;

class EnccryptedTest extends \PHPUnit_Framework_TestCase {

    /**
     * @expectedException laacz\XLSParser\XLSEncryptedException
     */
    function testPasswordToOpen()
    {
        new Book(file_get_contents(__DIR__ . '/xls/password-to-open.xls'));
    }

    function testPasswordToModify()
    {
        $book = new Book(file_get_contents(__DIR__ . '/xls/password-to-modify.xls'));
        $this->assertEquals('A', (string)$book[0][1][1]);
    }

    /**
     * @expectedException laacz\XLSParser\XLSEncryptedException
     */
    function testPasswordToBoth()
    {
        new Book(file_get_contents(__DIR__ . '/xls/password-to-open-and-modify.xls'));
    }

}
