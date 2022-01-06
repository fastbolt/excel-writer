<?php

namespace Sleussink\ExcelWriter\Tests;

use PHPUnit\Framework\TestCase;
use Sleussink\ExcelWriter\TableStyle;

class TableStyleTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $style = new TableStyle();
        $style->setDataRowStyle(['data row'])
            ->setHeaderStyle(['header'])
            ->setHeaderRowHeight(100);

        self::assertEquals('data row', $style->getDataRowStyle()[0], 'data row style');
        self::assertEquals('header', $style->getHeaderStyle()[0], 'header style');
        self::assertEquals(100, $style->getHeaderRowHeight(), 'row height');
    }
}
