<?php

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\ColumnSetting;
use Fastbolt\ExcelWriter\DataConverter;
use Fastbolt\ExcelWriter\ExcelGenerator;
use Fastbolt\ExcelWriter\SpreadSheetType;
use Fastbolt\ExcelWriter\TableStyle;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter;
use PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\ExcelGenerator
 */
class ExcelGeneratorTest extends TestCase
{
    private $spreadsheetType;
    private $converter;
    private array $mockedDependencies;

    public function setUp(): void
    {
        $this->spreadsheetType = $this->createMock(SpreadSheetType::class);
        $this->converter = $this->createMock(DataConverter::class);
        $this->mockedDependencies = [
            $this->spreadsheetType,
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
        $generator->expects(self::never())->method('applyStyle');

        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetWithColumns(): void
    {
        $spreadsheetType = new SpreadSheetType();
        $spreadsheetType
            ->setMaxColName('B')
            ->setColumns([
                new ColumnSetting('foo'),
                new ColumnSetting('bar'),
            ]);

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $spreadsheetType,
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
        $spreadsheetType = new SpreadSheetType();
        $spreadsheetType->setStyle(new TableStyle())
            ->setMaxColName('A');
        $spreadsheetType->setColumns([new ColumnSetting('foo')]);

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $spreadsheetType,
                $this->converter,
            ])
            ->onlyMethods([
                'applyHeaderStyle',
                'applyStyle',
                'saveFile',
                'applyColumnHeaders',
                'applyColumnFormat',
            ])
            ->getMock();
        $generator->expects(self::once())->method('applyHeaderStyle');
        $generator->expects(self::once())->method('applyStyle');

        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetWithContent(): void
    {
        $spreadsheetType = new SpreadSheetType();
        $spreadsheetType->setColumns([new ColumnSetting('foo')])
            ->setContent(['content']);
        $spreadsheetType->setMaxColName('A');

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $spreadsheetType,
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

    public function testGenerateSpreadSheetApplyingAll(): void
    {
        $col         = $this->createMock(ColumnDimension::class);
        $sheet       = $this->createMock(Worksheet::class);
        $spreadsheet = $this->createMock(Spreadsheet::class);
        $autoFilter  = $this->createMock(AutoFilter::class);
        $col        ->expects(self::once())->method('setAutoSize');
        $sheet      ->method('getColumnDimensions')->willReturn([$col]);
        $sheet      ->method('getAutoFilter')->willReturn($autoFilter);
        $spreadsheet->method('getActiveSheet')->willReturn($sheet);
        $autoFilter ->expects(self::once())->method('setRange')
                                           ->with("B1:C14");

        $spreadsheetType = new SpreadSheetType();
        $spreadsheetType
            ->setColumns([new ColumnSetting('')])
            ->setMaxColName('A')
            ->setStyle(new TableStyle())
            ->setContent(['content'])
            ->setSpreadsheet($spreadsheet)
            ->setAutoFilterRange("B1:C14");

        $generator = $this->getMockBuilder(ExcelGenerator::class)
            ->setConstructorArgs([
                $spreadsheetType,
                $this->converter,
            ])
            ->onlyMethods([
                'applyColumnHeaders',
                'applyColumnFormat',
                'applyHeaderStyle',
                'applyStyle',
                'applyContent',
                'saveFile',
            ])
            ->getMock();
        $generator->expects(self::once())->method('applyColumnHeaders');
        $generator->expects(self::once())->method('applyColumnFormat');
        $generator->expects(self::once())->method('applyHeaderStyle');
        $generator->expects(self::once())->method('applyContent');
        $generator->expects(self::once())->method('applyStyle');
        $generator->expects(self::once())->method('saveFile');
        $generator->generateSpreadsheet('url');
    }

    public function testGenerateSpreadSheetNoColumnError(): void
    {
        $this->expectException(\ArgumentCountError::class);

        $generator = new ExcelGenerator();
        $generator->generateSpreadsheet();
    }

    public function testSaveFile(): void
    {
        $this->spreadsheetType = new SpreadSheetType();
        $this->spreadsheetType->setSpreadsheet(new Spreadsheet());

        $generator = new ExcelGenerator(
            $this->spreadsheetType,
            $this->converter
        );

        $generator->saveFile('../test.xlsx');

        self::assertFileExists('../test.xlsx', 'file should have been created');
        unlink('../test.xlsx');
    }

    public function testSaveFileWhileNoURL(): void
    {
        $this->spreadsheetType = new SpreadSheetType();
        $this->spreadsheetType->setSpreadsheet(new Spreadsheet());

        $generator = new ExcelGenerator(
            $this->spreadsheetType,
            $this->converter
        );

        $generator->saveFile();
        $timestamp = substr((string)time(), 0, -1);
        $pattern = sys_get_temp_dir()."/spreadsheet ".$timestamp."*";
        //in case that the second changes between write and read, look for file that is 10 seconds older
        $fallBackPattern = sys_get_temp_dir()."/spreadsheet ".($timestamp - 1)."*";
        $file = glob($pattern)[0] ?? glob($fallBackPattern)[0];

        self::assertFileExists($file, 'file should have been created in '.sys_get_temp_dir());
        unlink($file);
    }

    public function testApplyContentWithEntity(): void
    {
        $content = [
            new SpreadSheetType() //example object
        ];

        $columns = [
            (new ColumnSetting('foo'))->setName('A'),
            (new ColumnSetting('foo'))->setName('B'),
        ];

        $worksheet = $this->getMockBuilder(Worksheet::class)
            ->onlyMethods(['setCellValueExplicit'])
            ->getMock();

        $this->spreadsheetType = $this->getMockBuilder(SpreadSheetType::class)
            ->onlyMethods([
                'getColumns',
                'getContentStartRow',
                'getSheet',
            ])
            ->getMock();
        $this->spreadsheetType->method('getColumns')->willReturn($columns);
        $this->spreadsheetType->method('getContentStartRow')->willReturn(2);
        $this->spreadsheetType->method('getSheet')->willReturn($worksheet);

        $worksheet->expects(self::exactly(2))
            ->method('setCellValueExplicit')
            ->withConsecutive(
                ['A2', 'foo'],
                ['B2', 'bar']
            );

        $converter = $this->getMockBuilder(DataConverter::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['convertEntityToArray'])
            ->getMock();
        $converter
            ->expects(self::once())
            ->method('convertEntityToArray')
            ->with($content, $columns)
            ->willReturn([['foo', 'bar']]);

        $generator = new ExcelGenerator(
            $this->spreadsheetType,
            $converter
        );

        $generator->applyContent($content);
    }

    public function testApplyContentWithArray(): void
    {
        $cols = [(new ColumnSetting('test', ColumnSetting::FORMAT_INTEGER))->setName('A')];
        $content = [['foo']];

        $this->converter->expects(self::once())->method('resolveCallableGetters')
            ->with($content, $cols)
            ->willReturn($content);

        $sheet = $this->createMock(Worksheet::class);
        $sheet->expects(self::once())
            ->method('setCellValue')
            ->with('A0', 'foo');
        $spreadsheetType = $this->createMock(SpreadSheetType::class);
        $spreadsheetType->method('getSheet')
            ->willReturn($sheet);
        $spreadsheetType->method('getColumns')
            ->willReturn($cols);

        $generator = new ExcelGenerator(
            $spreadsheetType,
            $this->converter
        );
        $generator->applyContent($content);
    }

    public function testApplyStyle(): void
    {
        $style = (new TableStyle())
            ->setHeaderRowHeight(100)
            ->setDataRowStyle(['data style']);

        $sheetStyle = $this->createMock(Style::class);
        $sheetStyle
            ->expects(self::once())
            ->method('applyFromArray')
            ->with(
                ['data style']
            )
            ->willReturn($sheetStyle);

        $sheet = $this->getMockBuilder(Worksheet::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['getStyle'])
            ->getMock();
        $sheet->method('getStyle')
            ->willReturn($sheetStyle);
        $spreadsheetType = $this->getMockBuilder(SpreadSheetType::class)
            ->onlyMethods(['getSheet'])
            ->getMock();
        $spreadsheetType->method('getSheet')
            ->willReturn($sheet);
        $spreadsheetType->setMaxColName('C')
            ->setMaxRowNumber(100);

        $generator = new ExcelGenerator(
            $spreadsheetType
        );

        $generator->applyStyle($style);
    }

    public function testApplyColumnHeaders(): void
    {
        $worksheet = $this->createMock(Worksheet::class);
        $spreadsheetType = $this->createMock(SpreadSheetType::class);
        $spreadsheetType->method('getSheet')->willReturn($worksheet);
        $worksheet->expects(self::exactly(2))
            ->method('setCellValue')
            ->withConsecutive(
                ['A1', 'foo'],
                ['B1', 'bar']
            );

        $columns = [
            new ColumnSetting('foo'),
            new ColumnSetting('bar'),
        ];

        $generator = new ExcelGenerator(
            $spreadsheetType
        );

        $generator->applyColumnHeaders($columns);
    }

    public function testApplyHeaderStyleNoHeaderHeight(): void
    {
        $this->spreadsheetType->expects(self::never())->method('getMaxColName');

        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $style = new TableStyle();
        $style->setHeaderRowHeight(0);

        $generator->applyHeaderStyle($style);
    }

    public function testApplyHeaderStyleMergeTwoRows(): void
    {
        $sheetStyle = $this->createMock(Style::class);
        $sheetStyle->method('applyFromArray')->willReturn($sheetStyle);
        $sheetStyle->expects(self::exactly(2))
            ->method('applyFromArray');

        $sheet = $this->createMock(Worksheet::class);
        $sheet->expects(self::exactly(3))
            ->method('mergeCells')
            ->withConsecutive(['A1:A2'], ['B1:B2'], ['C1:C2']);
        $sheet->method('getStyle')
            ->willReturn($sheetStyle);

        $this->spreadsheetType->method('getSheet')->willReturn($sheet);
        $this->spreadsheetType->method('getMaxColName')->willReturn('C');
        $this->spreadsheetType->method('getColumns')->willReturn(['A', 'B', 'C']);

        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $style = new TableStyle();
        $style->setHeaderRowHeight(2)
            ->setHeaderStyle([]);

        $generator->applyHeaderStyle($style);
    }

    public function testSetSpreadsheet(): void
    {
        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $spreadsheet = new Spreadsheet();

        $this->spreadsheetType->expects(self::once())
            ->method('setSpreadsheet')
            ->with($spreadsheet);

        $generator->setSpreadsheet($spreadsheet);
    }

    public function testSetColumns(): void
    {
        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $columns = [
            new ColumnSetting('First'),
            new ColumnSetting('Sec'),
            new ColumnSetting('Thrd'),
        ];

        $this->spreadsheetType->expects(self::once())
            ->method('setColumns')
            ->with($columns);
        $this->spreadsheetType->expects(self::once())
            ->method('setMaxColName')
            ->with('C');

        $generator->setColumns($columns);
    }

    public function testSetStyle(): void
    {
        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $style = new TableStyle();
        $this->spreadsheetType->expects(self::once())
            ->method('setStyle')
            ->with($style);

        $generator->setStyle($style);
    }

    public function testSetContent(): void
    {
        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $content = ['content'];
        $this->spreadsheetType->expects(self::once())
            ->method('setContent')
            ->with($content);

        $generator->setContent($content);
    }

    public function testSetAutoFilterRange(): void
    {
        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $this->spreadsheetType->expects(self::once())
            ->method('setAutoFilterRange')
            ->with($range = "foo");

        $generator->setAutoFilterRange($range);
    }

    public function testApplyColumnFormatNoNumberFormat(): void
    {
        $columns = [
            (new ColumnSetting('heading', ColumnSetting::FORMAT_INTEGER))->setName('A'),
        ];

        $style = $this->createMock(Style::class);
        $style->expects(self::once())->method('applyFromArray')
            ->with([
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_RIGHT,
                ],
                'numberFormat' => [
                    'formatCode' => NumberFormat::FORMAT_NUMBER,
                ],
            ]);

        $sheet = $this->createMock(Worksheet::class);
        $sheet->expects(self::once())->method('getStyle')->with('A:A')
            ->willReturn($style);

        $this->spreadsheetType->method('getSheet')->willReturn($sheet);

        $generator = new ExcelGenerator(
            $this->spreadsheetType
        );

        $generator->applyColumnFormat($columns);
    }
}
