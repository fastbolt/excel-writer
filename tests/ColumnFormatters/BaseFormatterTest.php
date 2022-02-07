<?php

namespace ColumnFormatters;

use Fastbolt\ExcelWriter\ColumnFormatter\BaseFormatter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ColumnFormatter\BaseFormatter
 */
class BaseFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new BaseFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
        self::assertNull($formatter->getNumberFormat(), 'number format');
    }
}
