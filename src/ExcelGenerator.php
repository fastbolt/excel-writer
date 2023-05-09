<?php

namespace Fastbolt\ExcelWriter;

use ArgumentCountError;
use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SplFileInfo;

class ExcelGenerator
{
    //TODO set the apply... and save functions to private, adapt tests
    private SpreadSheetType $spreadsheetType;
    private DataConverter $converter;

    public function __construct(
        ?SpreadSheetType $spreadsheetType = null,
        ?DataConverter $converter = null
    ) {
        $this->spreadsheetType  = $spreadsheetType ?? new SpreadSheetType();
        $this->converter        = $converter       ?? new DataConverter();
    }

    public function setSpreadsheet(Spreadsheet $spreadsheet): ExcelGenerator
    {
        $this->spreadsheetType->setSpreadsheet($spreadsheet);

        return $this;
    }

    /**
     * @param ColumnSetting[] $columns
     *
     * @return ExcelGenerator
     */
    public function setColumns(array $columns): ExcelGenerator
    {
        $this->spreadsheetType->setColumns($columns);

        $maxColName = Coordinate::stringFromColumnIndex(count($columns));
        $this->spreadsheetType->setMaxColName($maxColName);

        return $this;
    }

    public function setStyle(TableStyle $style): ExcelGenerator
    {
        $this->spreadsheetType->setStyle($style);

        return $this;
    }

    public function setContent(array $content): ExcelGenerator
    {
        $sheetType = $this->spreadsheetType;

        $sheetType->setContent($content);

        return $this;
    }

    /**
     * @param string $range
     * @return $this
     */
    public function setAutoFilterRange(string $range): ExcelGenerator
    {
        $this->spreadsheetType->setAutoFilterRange($range);

        return $this;
    }

    /**
     * @param array $cells
     * @return ExcelGenerator
     */
    public function mergeCells(array $cells): ExcelGenerator
    {
        $this->spreadsheetType->addMergedCells($cells);

        return $this;
    }

    /**
     * Generates a spreadsheet with a single sheet/table, using the previously set options
     *
     * @param string $url the path to where the file is supposed to be saved to (includes filename)
     *
     * @return SplFileInfo
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateSpreadsheet(string $url = ''): SplFileInfo
    {
        $headerRowHeight = $this->spreadsheetType->getStyle()->getHeaderRowHeight();
        $this->spreadsheetType
            ->setMaxRowNumber(count($this->spreadsheetType->getContent()) + $headerRowHeight)
            ->setContentStartRow($headerRowHeight + 1);

        if (count($this->spreadsheetType->getColumns()) === 0) {
            throw new ArgumentCountError('At least one column must be set.');
        }

        $this->applyColumnHeaders($this->spreadsheetType->getColumns());
        $this->applyColumnFormat($this->spreadsheetType->getColumns());
        $this->applyHeaderStyle($this->spreadsheetType->getStyle());

        if ($this->spreadsheetType->getContent()) {
            $this->applyContent($this->spreadsheetType->getContent());
        }

        $this->applyTableStyle($this->spreadsheetType->getStyle());
        $this->applyColumnStyle();

        //auto filter
        $sheet = $this->spreadsheetType->getSheet();
        if ($this->spreadsheetType->getAutoFilterRange() !== '') {
            $sheet->getAutoFilter()->setRange($this->spreadsheetType->getAutoFilterRange());
        }

        //auto size
        $dimensions = $sheet->getColumnDimensions();
        foreach ($dimensions as $col) {
            $col->setAutoSize(true);
        }

        $this->applyMergedCells();

        $file = $this->saveFile($url);

        //reset sheet to avoid leftover data when called multiple times
        $this->spreadsheetType->setSpreadsheet(new Spreadsheet());

        return $file;
    }

    /**
     * @param string $url
     *
     * @return SplFileInfo
     * @throws Exception
     */
    public function saveFile(string $url = ''): SplFileInfo
    {
        if ($url === '') {
            $url = sys_get_temp_dir() . '/spreadsheet ' . time();
        }

        if (!strpos($url, '.xlsx')) {
            $url .= '.xlsx';
        }

        $writer = new Xlsx($this->spreadsheetType->getSpreadsheet());
        $writer->save($url);

        return new SplFileInfo($url);
    }

    /**
     * @param array $content
     *
     * @return ExcelGenerator
     * @throws \Exception
     */
    public function applyContent(array $content): ExcelGenerator
    {
        $sheet      = $this->spreadsheetType->getSheet();
        $cols       = $this->spreadsheetType->getColumns();
        $currentRow = $this->spreadsheetType->getContentStartRow();
        $colCount   = count($cols);

        $content = array_values($content);
        if (getType($content[0]) === 'object') {
            //convert entities to usable arrays by calling their getters and callables
            $content = $this->converter->convertEntityToArray($content, $cols);
        } else {
            $content = $this->converter->resolveCallableGetters($content, $cols);
        }

        //write
        /** @var array $itemRow */
        foreach ($content as $itemRow) {
            $itemRow = array_values($itemRow);
            for ($counter = 0; $counter < $colCount; $counter++) {
                $col         = $cols[$counter];
                $colName     = $col->getName();
                $coordinates = $colName . $currentRow;

                //makes sure long numbers marked as string are displayed correctly
                if ($col->getFormatter() instanceof StringFormatter) {
                    $sheet->setCellValueExplicit($coordinates, $itemRow[$counter], DataType::TYPE_STRING);
                    continue;
                }

                $sheet->setCellValue($coordinates, $itemRow[$counter]);
            }
            $currentRow++;
        }

        return $this;
    }

    /**
     * @param TableStyle $style
     *
     * @return ExcelGenerator
     */
    public function applyTableStyle(TableStyle $style): ExcelGenerator
    {
        //set data row style
        $headerHeight     = $style->getHeaderRowHeight();
        $firstContentCell = 'A' . (1 + $headerHeight);
        $lastContentCell  = $this->spreadsheetType->getMaxColName() . $this->spreadsheetType->getMaxRowNumber();

        $this->spreadsheetType->getSheet()
            ->getStyle($firstContentCell . ':' . $lastContentCell)
            ->applyFromArray($style->getDataRowStyle());

        return $this;
    }

    /**
     * @return $this
     */
    public function applyColumnStyle(): ExcelGenerator
    {
        $spreadsheetType = $this->spreadsheetType;
        $columns         = $spreadsheetType->getColumns();
        $contentStartRow = $spreadsheetType->getContentStartRow();
        $headerHeight    = $spreadsheetType->getStyle()->getHeaderRowHeight();

        foreach ($columns as $col) {
            $colName = $col->getName();

            //header style
            if ($col->getHeaderStyle() !== null) {
                $spreadsheetType->getSheet()->getStyle(
                    $colName . "1:" . $colName . $headerHeight
                )->applyFromArray($col->getHeaderStyle());
            }

            //data row style
            if ($col->getDataStyle() !== null) {
                $spreadsheetType->getSheet()->getStyle(
                    $colName . $contentStartRow . ':'. $colName . $spreadsheetType->getMaxRowNumber()
                )->applyFromArray($col->getDataStyle());
            }
        }

        return $this;
    }

    /**
     * @param ColumnSetting[] $columns
     *
     * @return ExcelGenerator
     */
    public function applyColumnHeaders(array $columns): ExcelGenerator
    {
        $sheet       = $this->spreadsheetType->getSheet();
        $headerCount = count($columns);
        $letters  = range(1, $headerCount);
        array_walk($letters, static function (&$index) {
            $index = Coordinate::stringFromColumnIndex($index);
        });

        for ($counter = 0; $counter < $headerCount; $counter++) {
            $sheet->setCellValue($letters[$counter] . '1', $columns[$counter]->getHeader());
            $columns[$counter]->setName($letters[$counter]);
        }

        return $this;
    }

    /**
     * @param ColumnSetting[] $columns
     */
    public function applyColumnFormat(array $columns): void
    {
        $sheet = $this->spreadsheetType->getSheet();

        foreach ($columns as $column) {
            $format = [];
            $formatter = $column->getFormatter();

            $align          = $formatter->getAlignment();
            $numberFormat   = $formatter->getNumberFormat();

            $format['alignment'] = $align;

            if ($numberFormat) {
                $format['numberFormat'] = $numberFormat;
            }

            $sheet->getStyle($column->getName() . ':' . $column->getName())
                  ->applyFromArray($format);
        }
    }

    /**
     * @param TableStyle $style
     *
     * @return ExcelGenerator
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function applyHeaderStyle(TableStyle $style): ExcelGenerator
    {
        $sheet = $this->spreadsheetType->getSheet();

        if ($style->getHeaderRowHeight() < 1) {
            return $this;
        }

        //merging header row
        if ($style->getHeaderRowHeight() > 1) {
            $cols = range(1, count($this->spreadsheetType->getColumns()));
            array_walk($cols, static function (&$index) {
                $index = Coordinate::stringFromColumnIndex($index);
            });

            foreach ($cols as $col) {
                $firstCell = $col . '1';
                $lastCell  = $col . $style->getHeaderRowHeight();
                $sheet->mergeCells($firstCell . ':' . $lastCell);
            }
            unset($firstCell, $lastCell);
        }

        //set style
        $lastCell = $this->spreadsheetType->getMaxColName() . $style->getHeaderRowHeight();
        $sheet
            ->getStyle('A1' . ':' . $lastCell)
            ->applyFromArray([
                'alignment' => [
                    'vertical'   => Alignment::VERTICAL_CENTER,
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_MEDIUM,
                        'color'       => ['argb' => 'FF366092'],
                    ],
                ],
            ])
            ->applyFromArray($style->getHeaderStyle());

        return $this;
    }

    /**
     * Merges previously set cells in the worksheet
     *
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function applyMergedCells(): void
    {
        $sheet       = $this->spreadsheetType->getSheet();
        $mergedCells = $this->spreadsheetType->getMergedCells();

        foreach ($mergedCells as $cells) {
            $sheet->mergeCells($cells)
                  ->getStyle($cells)->getAlignment()->setHorizontal('center');
        }
    }
}
