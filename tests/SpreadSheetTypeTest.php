<?php

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\ColumnSetting;
use Fastbolt\ExcelWriter\SpreadSheetType;
use Fastbolt\ExcelWriter\TableStyle;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\SpreadSheetType
 */
class SpreadSheetTypeTest extends TestCase
{
    public function testSettersGetters(): void
    {
        $item = new SpreadSheetType();
        $item->setStyle(new TableStyle())
            ->setContent(['content'])
            ->setColumns([new ColumnSetting('')])
            ->setSpreadsheet($spreadsheet = new Spreadsheet())
            ->setMaxColName('foo')
            ->setMaxRowNumber(200)
            ->setContentStartRow(100);

        self::assertInstanceOf(TableStyle::class, $item->getStyle(), 'style');
        self::assertEquals('content', $item->getContent()[0], 'content');
        self::assertInstanceOf(ColumnSetting::class, $item->getColumns()[0], 'columns');
        self::assertSame($spreadsheet, $item->getSpreadsheet(), 'spreadsheet');
        self::assertNotNull($item->getSheet(), 'sheet');
        self::assertEquals('foo', $item->getMaxColName(), 'max col name');
        self::assertEquals(200, $item->getMaxRowNumber(), 'max row number');
        self::assertEquals(100, $item->getContentStartRow(), 'content start row');
    }
}
