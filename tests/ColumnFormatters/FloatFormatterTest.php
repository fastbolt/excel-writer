<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests\ColumnFormatters;

use Fastbolt\ExcelWriter\ColumnFormatter\FloatFormatter;
use Fastbolt\ExcelWriter\ColumnSetting;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ColumnFormatter\FloatFormatter
 */
class FloatFormatterTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $column = new ColumnSetting('foo', ColumnSetting::FORMAT_FLOAT, null, 3);
        $formatter = new FloatFormatter($column);

        self::assertEquals(['horizontal' => Alignment::HORIZONTAL_RIGHT], $formatter->getAlignment(), 'alignment');
        self::assertEquals(['formatCode' => 0.000], $formatter->getNumberFormat(), 'number format');
    }
}
