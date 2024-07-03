<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\ColumnSetting;
use Fastbolt\ExcelWriter\DataConverter;
use Fastbolt\ExcelWriter\ExcelGenerator;
use Fastbolt\ExcelWriter\TableStyle;
use Fastbolt\ExcelWriter\WorksheetType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ExcelGenerator
 */
class ExcelGeneratorTest extends TestCase
{
    private $worksheetType;
    private $converter;
    private array $mockedDependencies;

    public function setUp(): void
    {
        $this->worksheetType = $this->createMock(WorksheetType::class);
        $this->converter = $this->createMock(DataConverter::class);
        $this->mockedDependencies = [
            $this->worksheetType,
            $this->converter,
        ];
    }

    public function testGenerateSpreadsheetAllFalse(): void
    {
        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs($this->mockedDependencies)
            ->getMock();

        $generator->expects(self::never())->method('applyColumnHeaders');
        $generator->expects(self::never())->method('applyColumnFormat');
        $generator->expects(self::never())->method('applyHeaderStyle');
        $generator->expects(self::never())->method('applyContent');
        $generator->expects(self::never())->method('applyTableStyle');
        $generator->expects(self::never())->method('applyHeaderStyle');

        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetWithColumns(): void
    {
        $worksheetType = new WorksheetType(new Spreadsheet());
        $worksheetType
            ->setMaxColName('B')
            ->setColumns([
                new ColumnSetting('foo'),
                new ColumnSetting('bar'),
            ]);

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $worksheetType,
                $this->converter,
            ])
            ->onlyMethods(['applyColumnHeaders', 'applyColumnFormat', 'saveFile'])
            ->getMock();
        $generator->expects(self::once())->method('applyColumnHeaders');
        $generator->expects(self::once())->method('applyColumnFormat');

        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetWithStyle(): void
    {
        $worksheetType = new WorksheetType(new Spreadsheet());
        $worksheetType->setStyle(new TableStyle())
            ->setMaxColName('A');
        $worksheetType->setColumns([new ColumnSetting('foo')]);

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $worksheetType,
                $this->converter,
            ])
            ->onlyMethods([
                'applyHeaderStyle',
                'applyTableStyle',
                'applyColumnStyle',
                'saveFile',
                'applyColumnHeaders',
                'applyColumnFormat',
            ])
            ->getMock();
        $generator->expects(self::once())->method('applyHeaderStyle');
        $generator->expects(self::once())->method('applyTableStyle');
        $generator->expects(self::once())->method('applyColumnStyle');

        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetWithContent(): void
    {
        $worksheetType = new WorksheetType(new Spreadsheet());
        $worksheetType->setColumns([new ColumnSetting('foo')])
            ->setContent(['content']);
        $worksheetType->setMaxColName('A');

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $worksheetType,
                $this->converter,
            ])
            ->onlyMethods([
                'applyColumnHeaders',
                'applyColumnFormat',
                'applyContent',
            ])
            ->getMock();
        $generator->expects(self::once())->method('applyContent');

        $generator->generateSpreadsheet('url');
    }

//    public function testGenerateSpreadSheetApplyingAll(): void
//    {
//        $col = $this->createMock(ColumnDimension::class);
//        $sheet = $this->createMock(Worksheet::class);
//        $spreadsheet = $this->createMock(Spreadsheet::class);
//        $autoFilter = $this->createMock(AutoFilter::class);
//        $col->expects(self::once())->method('setAutoSize');
//        $sheet->method('getColumnDimensions')->willReturn([$col]);
//        $sheet->method('getAutoFilter')->willReturn($autoFilter);
//        $spreadsheet->method('getActiveSheet')->willReturn($sheet);
//        $autoFilter->expects(self::once())->method('setRange')
//            ->with("B1:C14");
//
//        $worksheetType = new WorksheetType(new Spreadsheet());
//        $worksheetType
//            ->setColumns([new ColumnSetting('')])
//            ->setMaxColName('A')
//            ->setStyle(new TableStyle())
//            ->setContent(['content'])
//            ->setSpreadsheet($spreadsheet)
//            ->setAutoFilterRange("B1:C14");
//
//        $generator = $this->getMockBuilder(ExcelGenerator::class)
//            ->setConstructorArgs([
//                $worksheetType,
//                $this->converter,
//            ])
//            ->onlyMethods([
//                'applyColumnHeaders',
//                'applyColumnFormat',
//                'applyHeaderStyle',
//                'applyTableStyle',
//                'applyColumnStyle',
//                'applyContent',
//                'saveFile',
//                'applyMergedCells'
//            ])
//            ->getMock();
//
//        $generator->expects(self::once())->method('applyColumnHeaders');
//        $generator->expects(self::once())->method('applyColumnFormat');
//        $generator->expects(self::once())->method('applyHeaderStyle');
//        $generator->expects(self::once())->method('applyContent');
//        $generator->expects(self::once())->method('applyTableStyle');
//        $generator->expects(self::once())->method('applyColumnStyle');
//        $generator->expects(self::once())->method('saveFile');
//        $generator->expects(self::once())->method('applyMergedCells');
//        $generator->generateSpreadsheet('url');
//    }

//    public function testGenerateSpreadSheetNoColumnError(): void
//    {
//        $this->expectException(\ArgumentCountError::class);
//
//        $generator = new ExcelGenerator();
//        $generator->generateSpreadsheet();
//    }

    public function testSaveFile(): void
    {
        $this->worksheetType = new WorksheetType(new Spreadsheet());
        $this->worksheetType->setSpreadsheet(new Spreadsheet());

        $generator = new ExcelGenerator(
            $this->worksheetType,
            $this->converter
        );

        $generator->saveFile('../test.xlsx');

        self::assertFileExists('../test.xlsx', 'file should have been created');
        unlink('../test.xlsx');
    }

    public function testSaveFileWhileNoURL(): void
    {
        $this->worksheetType = new WorksheetType(new Spreadsheet());
        $this->worksheetType->setSpreadsheet(new Spreadsheet());

        $generator = new ExcelGenerator(
            $this->worksheetType,
            $this->converter
        );

        $generator->saveFile();
        $timestamp = substr((string)time(), 0, -1);
        $pattern = sys_get_temp_dir() . "/spreadsheet " . $timestamp . "*";
        //in case that the second changes between write and read, look for file that is 10 seconds older
        $fallBackPattern = sys_get_temp_dir() . "/spreadsheet " . ($timestamp - 1) . "*";
        $file = glob($pattern)[0] ?? glob($fallBackPattern)[0];

        self::assertFileExists($file, 'file should have been created in ' . sys_get_temp_dir());
        unlink($file);
    }

//    public function testApplyContentWithEntity(): void
//    {
//        $content = [
//            new WorksheetType(new Spreadsheet()), //example object
//        ];
//
//        $columns = [
//            (new ColumnSetting('foo'))->setName('A'),
//            (new ColumnSetting('foo'))->setName('B'),
//        ];
//
//        $worksheet = $this->getMockBuilder(Worksheet::class)
//            ->onlyMethods(['setCellValueExplicit'])
//            ->getMock();
//
//        $this->worksheetType = $this->getMockBuilder(WorksheetType::class)
//            ->onlyMethods([
//                'getColumns',
//                'getContentStartRow',
//                'getSpreadsheet',
//            ])
//            ->getMock();
//        $this->worksheetType->method('getColumns')->willReturn($columns);
//        $this->worksheetType->method('getContentStartRow')->willReturn(2);
//        $this->worksheetType->method('getSpreadsheet')->willReturn($worksheet);
//
//        $worksheet->expects(self::exactly(2))
//            ->method('setCellValueExplicit')
//            ->withConsecutive(
//                ['A2', 'foo'],
//                ['B2', 'bar']
//            );
//
//        $converter = $this->getMockBuilder(DataConverter::class)
//            ->disableOriginalConstructor()
//            ->onlyMethods(['convertEntityToArray'])
//            ->getMock();
//        $converter
//            ->expects(self::once())
//            ->method('convertEntityToArray')
//            ->with($content, $columns)
//            ->willReturn([['foo', 'bar']]);
//
//        $generator = new ExcelGenerator(
//            $this->worksheetType,
//            $converter
//        );
//
//        $generator->applyContent($content);
//    }

//    public function testApplyContentWithArray(): void
//    {
//        $cols = [(new ColumnSetting('test', ColumnSetting::FORMAT_INTEGER))->setName('A')];
//        $content = [['foo']];
//
//        $this->converter->expects(self::once())->method('resolveCallableGetters')
//            ->with($content, $cols)
//            ->willReturn($content);
//
//        $sheet = $this->createMock(Worksheet::class);
//        $sheet->expects(self::once())
//            ->method('setCellValue')
//            ->with('A0', 'foo');
//        $worksheetType = $this->createMock(WorksheetType::class);
//        $worksheetType->method('getSpreadsheet')
//            ->willReturn($sheet);
//        $worksheetType->method('getColumns')
//            ->willReturn($cols);
//
//        $generator = new ExcelGenerator(
//            $worksheetType,
//            $this->converter
//        );
//        $generator->applyContent($content);
//    }

//    public function testApplyTableStyle(): void
//    {
//        $style = (new TableStyle())
//            ->setHeaderRowHeight(100)
//            ->setDataRowStyle(['data style']);
//
//        $sheetStyle = $this->createMock(Style::class);
//        $sheetStyle
//            ->expects(self::once())
//            ->method('applyFromArray')
//            ->with(
//                ['data style']
//            )
//            ->willReturn($sheetStyle);
//
//        $sheet = $this->getMockBuilder(Worksheet::class)
//            ->disableOriginalConstructor()
//            ->onlyMethods(['getStyle'])
//            ->getMock();
//        $sheet->method('getStyle')
//            ->willReturn($sheetStyle);
//        $worksheetTpye = $this->getMockBuilder(WorksheetType::class)
//            ->onlyMethods(['getSpreadsheet'])
//            ->getMock();
//        $worksheetTpye->method('getSpreadsheet')
//            ->willReturn($sheet);
//        $worksheetTpye->setMaxColName('C')
//            ->setMaxRowNumber(100);
//
//        $generator = new ExcelGenerator(
//            $worksheetTpye
//        );
//
//        $generator->applyTableStyle($style);
//    }

//    public function testApplyColumnStyle(): void
//    {
//        $sheet = $this->createMock(Worksheet::class);
//        $spreadsheet = $this->getMockBuilder(Spreadsheet::class)
//            ->disableOriginalConstructor()
//            ->onlyMethods(['getActiveSheet'])
//            ->getMock();
//        $spreadsheet->method('getActiveSheet')
//            ->willReturn($sheet);
//        $column = (new ColumnSetting('Foo'))
//            ->setName('B')
//            ->setHeaderStyle(['headerStyle'])
//            ->setDataStyle(['dataStyle']);
//        $tableStyle = (new TableStyle())->setHeaderRowHeight(4);
//        $worksheetType = (new WorksheetType(new Spreadsheet()))
//            ->setStyle($tableStyle)
//            ->setContentStartRow(5)
//            ->setMaxRowNumber(10)
//            ->setColumns([$column])
//            ->setSpreadsheet($spreadsheet);
//
//        $generator = new ExcelGenerator($worksheetType);
//        $sheetStyle = $this->createMock(Style::class);
//
//        $sheet->expects(self::exactly(2))
//            ->method('getStyle')
//            ->withConsecutive(['B1:B4'], ['B5:B10'])
//            ->willReturn($sheetStyle);
//
//        $sheetStyle->expects(self::exactly(2))
//            ->method('applyFromArray')
//            ->withConsecutive(
//                [['headerStyle'], true],
//                [['dataStyle'], true]
//            );
//
//        $generator->applyColumnStyle();
//    }

//    public function testApplyColumnStyleNoColumns(): void
//    {
//        $worksheetType = $this->createMock(WorksheetType::class);
//        $worksheetType->expects(self::never())->method('getSpreadsheet');
//        $generator = new ExcelGenerator($worksheetType);
//        $generator->applyColumnStyle();
//    }
//
//    public function testApplyColumnStyleWhenHeaderStyleOnly(): void
//    {
//        $sheet = $this->createMock(Worksheet::class);
//        $spreadsheet = $this->getMockBuilder(Spreadsheet::class)
//            ->disableOriginalConstructor()
//            ->onlyMethods(['getActiveSheet'])
//            ->getMock();
//        $spreadsheet->method('getActiveSheet')
//            ->willReturn($sheet);
//        $column = (new ColumnSetting('Foo'))
//            ->setName('B')
//            ->setHeaderStyle(['headerStyle']);
//        $tableStyle = (new TableStyle())->setHeaderRowHeight(4);
//        $worksheetType = (new WorksheetType(new Spreadsheet()))
//            ->setStyle($tableStyle)
//            ->setContentStartRow(5)
//            ->setMaxRowNumber(10)
//            ->setColumns([$column])
//            ->setSpreadsheet($spreadsheet);
//
//        $generator = new ExcelGenerator($worksheetType);
//        $sheetStyle = $this->createMock(Style::class);
//
//        $sheet->expects(self::once())
//            ->method('getStyle')
//            ->with('B1:B4')
//            ->willReturn($sheetStyle);
//
//        $sheetStyle->expects(self::once())
//            ->method('applyFromArray')
//            ->with(['headerStyle'], true);
//
//        $generator->applyColumnStyle();
//    }
//
//    public function testApplyColumnStyleWhenDataStyleOnly(): void
//    {
//        $sheet = $this->createMock(Worksheet::class);
//        $spreadsheet = $this->getMockBuilder(Spreadsheet::class)
//            ->disableOriginalConstructor()
//            ->onlyMethods(['getActiveSheet'])
//            ->getMock();
//        $spreadsheet->method('getActiveSheet')
//            ->willReturn($sheet);
//        $column = (new ColumnSetting('Foo'))
//            ->setName('B')
//            ->setDataStyle(['dataStyle']);
//        $tableStyle = (new TableStyle())->setHeaderRowHeight(4);
//        $worksheetType = (new WorksheetType(new Spreadsheet()))
//            ->setStyle($tableStyle)
//            ->setContentStartRow(5)
//            ->setMaxRowNumber(10)
//            ->setColumns([$column])
//            ->setSpreadsheet($spreadsheet);
//
//        $generator = new ExcelGenerator($worksheetType);
//        $sheetStyle = $this->createMock(Style::class);
//
//        $sheet->expects(self::once())
//            ->method('getStyle')
//            ->with('B5:B10')
//            ->willReturn($sheetStyle);
//
//        $sheetStyle->expects(self::once())
//            ->method('applyFromArray')
//            ->with(['dataStyle'], true);
//
//        $generator->applyColumnStyle();
//    }
//
//    public function testApplyColumnHeaders(): void
//    {
//        $worksheet = $this->createMock(Worksheet::class);
//        $worksheetType = $this->createMock(WorksheetType::class);
//        $worksheetType->method('getSpreadsheet')->willReturn($worksheet);
//        $worksheet->expects(self::exactly(2))
//            ->method('setCellValue')
//            ->withConsecutive(
//                ['A1', 'foo'],
//                ['B1', 'bar']
//            );
//
//        $columns = [
//            new ColumnSetting('foo'),
//            new ColumnSetting('bar'),
//        ];
//
//        $generator = new ExcelGenerator(
//            $worksheetType
//        );
//
//        $generator->applyColumnHeaders($columns);
//    }
//
//    public function testApplyHeaderStyleNoHeaderHeight(): void
//    {
//        $this->worksheetType->expects(self::never())->method('getMaxColName');
//
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $style = new TableStyle();
//        $style->setHeaderRowHeight(0);
//
//        $generator->applyHeaderStyle($style);
//    }
//
//    public function testApplyHeaderStyleMergeTwoRows(): void
//    {
//        $sheetStyle = $this->createMock(Style::class);
//        $sheetStyle->method('applyFromArray')->willReturn($sheetStyle);
//        $sheetStyle->expects(self::exactly(2))
//            ->method('applyFromArray');
//
//        $sheet = $this->createMock(Worksheet::class);
//        $sheet->expects(self::exactly(3))
//            ->method('mergeCells')
//            ->withConsecutive(['A1:A2'], ['B1:B2'], ['C1:C2']);
//        $sheet->method('getStyle')
//            ->willReturn($sheetStyle);
//
//        $this->worksheetType->method('getSpreadsheet')->willReturn($sheet);
//        $this->worksheetType->method('getMaxColName')->willReturn('C');
//        $this->worksheetType->method('getColumns')->willReturn(['A', 'B', 'C']);
//
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $style = new TableStyle();
//        $style->setHeaderRowHeight(2)
//            ->setHeaderStyle([]);
//
//        $generator->applyHeaderStyle($style);
//    }
//
//    public function testSetSpreadsheet(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $spreadsheet = new Spreadsheet();
//
//        $this->worksheetType->expects(self::once())
//            ->method('setSpreadsheet')
//            ->with($spreadsheet);
//
//        $generator->setSpreadsheet($spreadsheet);
//    }
//
//    public function testSetColumns(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $columns = [
//            new ColumnSetting('First'),
//            new ColumnSetting('Sec'),
//            new ColumnSetting('Thrd'),
//        ];
//
//        $this->worksheetType->expects(self::once())
//            ->method('setColumns')
//            ->with($columns);
//        $this->worksheetType->expects(self::once())
//            ->method('setMaxColName')
//            ->with('C');
//
//        $generator->setColumns($columns);
//    }
//
//    public function testSetStyle(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $style = new TableStyle();
//        $this->worksheetType->expects(self::once())
//            ->method('setStyle')
//            ->with($style);
//
//        $generator->setStyle($style);
//    }
//
//    public function testSetContent(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $content = ['content'];
//        $this->worksheetType->expects(self::once())
//            ->method('setContent')
//            ->with($content);
//
//        $generator->setContent($content);
//    }
//
//    public function testSetAutoFilterRange(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $this->worksheetType->expects(self::once())
//            ->method('setAutoFilterRange')
//            ->with($range = "foo");
//
//        $generator->setAutoFilterRange($range);
//    }
//
//    public function testMergeCells(): void
//    {
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $this->worksheetType->expects(self::once())
//            ->method('addMergedCells');
//        $generator->mergeCells(['foo']);
//    }
//
//    public function testApplyColumnFormatNoNumberFormat(): void
//    {
//        $columns = [
//            (new ColumnSetting('heading', ColumnSetting::FORMAT_INTEGER))->setName('A'),
//        ];
//
//        $style = $this->createMock(Style::class);
//        $style->expects(self::once())->method('applyFromArray')
//            ->with([
//                'alignment' => [
//                    'horizontal' => Alignment::HORIZONTAL_RIGHT,
//                ],
//                'numberFormat' => [
//                    'formatCode' => NumberFormat::FORMAT_NUMBER,
//                ],
//            ]);
//
//        $sheet = $this->createMock(Worksheet::class);
//        $sheet->expects(self::once())->method('getStyle')->with('A:A')
//            ->willReturn($style);
//
//        $this->worksheetType->method('getSpreadsheet')->willReturn($sheet);
//
//        $generator = new ExcelGenerator(
//            $this->worksheetType
//        );
//
//        $generator->applyColumnFormat($columns);
//    }
//
//    public function testApplyMergedCells(): void
//    {
//        $worksheetType = new WorksheetType(new Spreadsheet());
//        $generator = new ExcelGenerator($worksheetType);
//        $spreadsheet = $this->createMock(Spreadsheet::class);
//        $sheet = $this->getMockBuilder(Worksheet::class)
//            ->onlyMethods(['mergeCells', 'getStyle'])
//            ->addMethods(['getAlignment', 'setHorizontal'])
//            ->getMock();
//        $sheet->method('getAlignment')->willReturn($sheet);
//
//        $worksheetType->setSpreadsheet($spreadsheet)
//            ->setMergedCells(['foo', 'bah', 'ham']);
//        $spreadsheet->method('getActiveSheet')->willReturn($sheet);
//
//        $sheet->expects(self::exactly(3))
//            ->method('mergeCells')
//            ->withConsecutive(['foo'], ['bah'], ['ham'])
//            ->willReturn($sheet);
//        $sheet->expects(self::exactly(3))
//            ->method('getStyle')
//            ->withConsecutive(['foo'], ['bah'], ['ham'])
//            ->willReturn($sheet);
//
//        $generator->applyMergedCells();
//    }
}
