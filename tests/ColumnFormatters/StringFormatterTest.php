<?php

namespace ColumnFormatters;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;
use Sleussink\ExcelWriter\ColumnFormatter\StringFormatter;

class StringFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new StringFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
    }
}
