<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests\ColumnFormatters;

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
        self::assertEquals(['horizontal' => Alignment::HORIZONTAL_LEFT], $formatter->getAlignment(), 'alignment');
        self::assertNull($formatter->getNumberFormat(), 'number format');
    }
}
