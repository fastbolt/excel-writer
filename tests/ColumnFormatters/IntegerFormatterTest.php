<?php

namespace ColumnFormatters;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;
use Sleussink\ExcelWriter\ColumnFormatter\IntegerFormatter;

class IntegerFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new IntegerFormatter();
        self::assertEquals(['vertical' => Alignment::HORIZONTAL_RIGHT], $formatter->getAlignment(), 'alignment');
        self::assertEquals(['formatCode' => 0.000], $formatter->getNumberFormat(), 'number format');
    }
}