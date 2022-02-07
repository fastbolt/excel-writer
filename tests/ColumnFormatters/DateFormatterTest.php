<?php

namespace ColumnFormatters;

use Fastbolt\ExcelWriter\ColumnFormatter\DateFormatter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ColumnFormatter\DateFormatter
 */
class DateFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new DateFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
    }
}
