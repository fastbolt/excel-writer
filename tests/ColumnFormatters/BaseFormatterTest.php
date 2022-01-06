<?php

namespace ColumnFormatters;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;
use Sleussink\ExcelWriter\ColumnFormatter\BaseFormatter;

class BaseFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new BaseFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
        self::assertNull($formatter->getNumberFormat(), 'number format');
    }
}
