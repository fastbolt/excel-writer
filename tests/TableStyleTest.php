<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\TableStyle;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\TableStyle
 */
class TableStyleTest extends TestCase
{
    public function testGetterSetter(): void
    {
        $style = new TableStyle();
        $style->setDataRowStyle(['data row'])
            ->setHeaderStyle(['header'])
            ->setHeaderRowHeight(100);

        self::assertEquals('data row', $style->getDataRowStyle()[0], 'data row style');
        self::assertEquals('header', $style->getHeaderStyle()[0], 'header style');
        self::assertEquals(100, $style->getHeaderRowHeight(), 'row height');
    }
}
