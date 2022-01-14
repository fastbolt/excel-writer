<?php

namespace ColumnFormatters;

use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;

class StringFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new StringFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
    }
}
