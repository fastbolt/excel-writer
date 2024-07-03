<?php

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\ColumnFormatter\DateFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\FloatFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\IntegerFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ColumnSetting
 */
class ColumnSettingTest extends TestCase
{
    public function testSettersGetters(): void
    {
        $column = new ColumnSetting('header', 'foo', 'getter', 4);
        self::assertEquals('header', $column->getHeader(), 'constructor header');
        self::assertInstanceOf(StringFormatter::class, $column->getFormatter(), 'constructor format');
        self::assertEquals('getter', $column->getGetter(), 'constructor getterName');
        self::assertEquals(4, $column->getDecimalLength(), 'constructor decimal length');

        $column->setName('foo')
               ->setGetter('foobar')
               ->setHeader('barfoo')
               ->setDecimalLength(5)
               ->setHeaderStyle(['hs'])
               ->setDataStyle(['ds']);
        self::assertEquals('foo', $column->getName(), 'get name');
        self::assertEquals('foobar', $column->getGetter(), 'get getterName');
        self::assertEquals('barfoo', $column->getHeader(), 'get header');
        self::assertEquals(5, $column->getDecimalLength(), 'decimal length');
        self::assertEquals(['hs'], $column->getHeaderStyle(), 'header style');
        self::assertEquals(['ds'], $column->getDataStyle(), 'data style');

        $column->setGetter(static function () {
            return 'ham';
        });
        self::assertEquals('ham', $column->getGetter()());
    }

    public function testGetFormatter(): void
    {
        $column = new ColumnSetting('test', ColumnSetting::FORMAT_FLOAT);
        self::assertInstanceOf(FloatFormatter::class, $column->getFormatter(), 'get format');

        $column->setFormat(ColumnSetting::FORMAT_STRING);
        self::assertInstanceOf(StringFormatter::class, $column->getFormatter(), 'get format');

        $column->setFormat(ColumnSetting::FORMAT_DATE);
        self::assertInstanceOf(DateFormatter::class, $column->getFormatter(), 'get format');

        $column->setFormat(ColumnSetting::FORMAT_INTEGER);
        self::assertInstanceOf(IntegerFormatter::class, $column->getFormatter(), 'get format');
    }
}
