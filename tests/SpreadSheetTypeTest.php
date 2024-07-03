<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\ColumnSetting;
use Fastbolt\ExcelWriter\TableStyle;
use Fastbolt\ExcelWriter\WorksheetType;
use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\WorksheetType
 */
class SpreadSheetTypeTest extends TestCase
{
    public function testSettersGetters(): void
    {
        $item = new WorksheetType($spreadsheet = new Spreadsheet());
        $item->setStyle(new TableStyle())
            ->setContent(['content'])
            ->setColumns([new ColumnSetting('')])
            ->setMaxColName('foo')
            ->setMaxRowNumber(200)
            ->setContentStartRow(100)
            ->setAutoFilterRange('B2:R15');

        self::assertInstanceOf(TableStyle::class, $item->getStyle(), 'style');
        self::assertEquals('content', $item->getContent()[0], 'content');
        self::assertInstanceOf(ColumnSetting::class, $item->getColumns()[0], 'columns');
        self::assertSame($spreadsheet, $item->getSpreadsheet(), 'spreadsheet');
        self::assertEquals('foo', $item->getMaxColName(), 'max col name');
        self::assertEquals(200, $item->getMaxRowNumber(), 'max row number');
        self::assertEquals(100, $item->getContentStartRow(), 'content start row');
        self::assertEquals('B2:R15', $item->getAutoFilterRange(), 'auto filter range');
    }

    public function testSetAutoFilterRangeNoRangeError(): void
    {
        $this->expectException(InvalidArgumentException::class);
        $item = new WorksheetType(new Spreadsheet());
        $item->setAutoFilterRange('A');
    }

    public function testSetAndAddMergeCells(): void
    {
        $item = new WorksheetType(new Spreadsheet());
        $item->setMergedCells(['foo', 'bar'])
             ->addMergedCells(['ham', 'eggs']);

        self::assertEquals(['foo', 'bar', 'ham', 'eggs'], $item->getMergedCells(), 'merged cells');
    }
}
