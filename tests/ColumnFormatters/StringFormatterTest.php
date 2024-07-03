<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests\ColumnFormatters;

use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter
 */
class StringFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $formatter = new StringFormatter();
        self::assertEquals(['horizontal' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
        self::assertEquals(['formatCode' => '#'], $formatter->getNumberFormat(), 'number format');
    }
}
