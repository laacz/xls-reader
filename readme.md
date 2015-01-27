*XLSParser* is reasonably fast PHP library intended to parse Microsoft Excel legacy binary XLS formats. It was written 
because all PHP implementations where too slow. Code is more or less direct port of python's excellent
[xlrd package](https://pypi.python.org/pypi/xlrd). Library is very early alpha. I made it a long time ago.

# Feedback

Issues and pull requests are accepted.

# Requirements

* PHP 5.4.0 or newer.
* [Multibyte string extension](http://php.net/mbstring) (*mbstring*) to handle UTF-16LE encoding, used in
newer Excel files.
* Little endian system because of PHP's unpack/pack. If you're not on *Sparc*, you should be covered.

# Install

Via command line: `composer require laacz/xls-parser`.

# Tests

Install dependencies with `composer install --dev`, then run tests with `vendor/bin/phpunit`.

# Usage

KISS. Provide filename and it gets loaded or parsed.

```php
$book = new laacz\XLSReader\Book(file_get_content('workbook.xls'));
```

# Accessing sheets

Sheets can be accessed via their numeric index or name. Since `Sheet` object implements *ArrayAccess* and *IteratorAggregate*, you
can do that too.

```php
$sheet = $book[0];
$sheet = $book['Page1'];
$sheet = $book['Vājprāts'];
```

# Accessing cells

Cells also can be accessed as with sheets. Index starts from zero.

```php
$row = $sheet[0];
$cell = $sheet[$sheet->nrows - 1][1];
```

To get value of a cell, cast it to string (or use it in such context) or get `value` attribute:

```php
$val1 = $cell->value;
$val2 = (string)$cell;
```

Or, if you wish...

```php
$val1 = $book[0][0][0]->value;
```

# Formatting

By now formatting can be accessed raw. In short - sheet contains mapping array `rich_text_runlist_map[][]`, which has 
arrays with two elements - position and font reference. First is position where style is being applied from, second is
number which refers to book's `font_list[]`, which on its part contains format description.

# Excel dates

Library does its best to parse dates found within cells. It returns string in common date format: 'yyyy-mm-dd hh:mi:ss'. 
For example: 2014-12-31 12:59:59.

# Wishlist

* Memory efficiency does not exist in context of this library.
* Performance might be better.
* Abstract formatting.
* Add helper methods for common tasks - returning columns, ranges, etc.
