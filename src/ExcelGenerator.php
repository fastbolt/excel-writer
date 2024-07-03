<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter;

use ArgumentCountError;
use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SplFileInfo;

class ExcelGenerator
{
    //TODO set the apply... and save functions to private, adapt tests

    private Spreadsheet $spreadsheet;

    /**
     * @var WorksheetType[] $worksheetTypes
     */
    private array $worksheetTypes = [];
    private DataConverter $converter;

    private WorksheetType $worksheetType;

    public function __construct(?WorksheetType $worksheetType = null, ?DataConverter $converter = null)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->worksheetType = $worksheetType ?? new WorksheetType($this->spreadsheet);
        $this->converter = $converter ?? new DataConverter();
    }

    public function setSpreadsheet(Spreadsheet $spreadsheet): ExcelGenerator
    {
        $this->spreadsheet = $spreadsheet;

        return $this;
    }

    /**
     * @param ColumnSetting[] $columns
     *
     * @return ExcelGenerator
     */
    public function setColumns(array $columns): ExcelGenerator
    {
        $this->worksheetType->setColumns($columns);

        $maxColName = Coordinate::stringFromColumnIndex(count($columns));
        $this->worksheetType->setMaxColName($maxColName);

        return $this;
    }

    public function setStyle(TableStyle $style): ExcelGenerator
    {
        $this->worksheetType->setStyle($style);

        return $this;
    }

    public function setContent(array $content): ExcelGenerator
    {
        $sheetType = $this->worksheetType;

        $sheetType->setContent($content);

        return $this;
    }

    /**
     * @param string $range
     * @return $this
     */
    public function setAutoFilterRange(string $range): ExcelGenerator
    {
        $this->worksheetType->setAutoFilterRange($range);

        return $this;
    }

    /**
     * @param array $cells
     * @return ExcelGenerator
     */
    public function mergeCells(array $cells): ExcelGenerator
    {
        $this->worksheetType->addMergedCells($cells);

        return $this;
    }

    public function setTitle(string $title): ExcelGenerator
    {
        $this->worksheetType->setTitle($title);

        return $this;
    }

    public function nextWorksheet(): self
    {
        $this->worksheetTypes[] = $this->worksheetType;

        $this->worksheetType = new WorksheetType($this->spreadsheet);

        return $this;
    }

    /**
     * @param string $url the path to where the file is supposed to be saved to (includes filename)
     *
     * @return SplFileInfo
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function generateSpreadsheet(string $url = ''): SplFileInfo
    {
        if (count($this->worksheetType->getColumns()) !== 0) {
            $this->worksheetTypes[] = $this->worksheetType;
        }
        $this->worksheetType = null; //all worksheets should be in worksheet array now

        $this->spreadsheet->removeSheetByIndex(0); //remove empty preset sheet
        foreach ($this->worksheetTypes as $index => $this->worksheetType) {
            $worksheet = new Worksheet($this->spreadsheet, $this->worksheetType->getTitle());
            $this->spreadsheet->addSheet($worksheet, $index);

            $this->worksheetType->setWorksheet($worksheet);
            $headerRowHeight = $this->worksheetType->getStyle()->getHeaderRowHeight();
            $this->worksheetType
                ->setMaxRowNumber(count($this->worksheetType->getContent()) + $headerRowHeight)
                ->setContentStartRow($headerRowHeight + 1);

            if (count($this->worksheetType->getColumns()) === 0) {
                throw new ArgumentCountError('At least one column must be set.');
            }

            $this->applyColumnHeaders();
            $this->applyColumnFormat();
            $this->applyHeaderStyle();
//
            if ($this->worksheetType->getContent()) {
                $this->applyContent($this->worksheetType->getContent());
            }
//
            $this->applyTableStyle();
            $this->applyColumnStyle();

            //auto filter
            if ($this->worksheetType->getAutoFilterRange() !== '') {
                $worksheet->getAutoFilter()->setRange($this->worksheetType->getAutoFilterRange());
            }

            //auto size
            $dimensions = $worksheet->getColumnDimensions();
            foreach ($dimensions as $col) {
                $col->setAutoSize(true);
            }

            $this->applyMergedCells();
        }

        $file = $this->saveFile($url);

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

        $writer = new Xlsx($this->spreadsheet);
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
        $sheet = $this->worksheetType->getWorksheet();
        $cols = $this->worksheetType->getColumns();
        $currentRow = $this->worksheetType->getContentStartRow();
        $colCount = count($cols);

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
                $col = $cols[$counter];
                $colName = $col->getName();

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
     * @return ExcelGenerator
     */
    public function applyTableStyle(): ExcelGenerator
    {
        $style = $this->worksheetType->getStyle();

        //set data row style
        $headerHeight = $style->getHeaderRowHeight();
        $firstContentCell = 'A' . (1 + $headerHeight);
        $lastContentCell = $this->worksheetType->getMaxColName() . $this->worksheetType->getMaxRowNumber();

        $this->worksheetType->getWorksheet()
            ->getStyle($firstContentCell . ':' . $lastContentCell)
            ->applyFromArray($style->getDataRowStyle());

        return $this;
    }

    /**
     * @return $this
     */
    public function applyColumnStyle(): ExcelGenerator
    {
        $columns = $this->worksheetType->getColumns();
        $contentStartRow = $this->worksheetType->getContentStartRow();
        $headerHeight = $this->worksheetType->getStyle()->getHeaderRowHeight();

        foreach ($columns as $col) {
            $colName = $col->getName();

            //header style
            if ($col->getHeaderStyle() !== null) {
                $this->worksheetType->getWorksheet()->getStyle(
                    $colName . "1:" . $colName . $headerHeight
                )->applyFromArray($col->getHeaderStyle());
            }

            //data row style
            if ($col->getDataStyle() !== null) {
                $this->worksheetType->getWorksheet()->getStyle(
                    $colName . $contentStartRow . ':' . $colName . $this->worksheetType->getMaxRowNumber()
                )->applyFromArray($col->getDataStyle());
            }
        }

        return $this;
    }

    /**
     * @return ExcelGenerator
     */
    public function applyColumnHeaders(): ExcelGenerator
    {
        $columns = $this->worksheetType->getColumns();
        $sheet = $this->worksheetType->getWorksheet();
        $headerCount = count($columns);
        $letters = range(1, $headerCount);
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
    public function applyColumnFormat(): void
    {
        $columns = $this->worksheetType->getColumns();
        $sheet = $this->worksheetType->getWorksheet();

        foreach ($columns as $column) {
            $format = [];
            $formatter = $column->getFormatter();

            $align = $formatter->getAlignment();
            $numberFormat = $formatter->getNumberFormat();

            $format['alignment'] = $align;

            if ($numberFormat) {
                $format['numberFormat'] = $numberFormat;
            }

            $sheet->getStyle($column->getName() . ':' . $column->getName())
                ->applyFromArray($format);
        }
    }

    /**
     *
     * @return ExcelGenerator
     */
    public function applyHeaderStyle(): ExcelGenerator
    {
        $style = $this->worksheetType->getStyle();
        $sheet = $this->worksheetType->getWorksheet();

        if ($style->getHeaderRowHeight() < 1) {
            return $this;
        }

        //merging header row
        if ($style->getHeaderRowHeight() > 1) {
            $cols = range(1, count($this->worksheetType->getColumns()));
            array_walk($cols, static function (&$index) {
                $index = Coordinate::stringFromColumnIndex($index);
            });

            foreach ($cols as $col) {
                $firstCell = $col . '1';
                $lastCell = $col . $style->getHeaderRowHeight();
                $sheet->mergeCells($firstCell . ':' . $lastCell);
            }
            unset($firstCell, $lastCell);
        }

        //set style
        $lastCell = $this->worksheetType->getMaxColName() . $style->getHeaderRowHeight();
        $sheet
            ->getStyle('A1' . ':' . $lastCell)
            ->applyFromArray([
                'alignment' => [
                    'vertical' => Alignment::VERTICAL_CENTER,
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_MEDIUM,
                        'color' => ['argb' => 'FF366092'],
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
        $sheet = $this->worksheetType->getWorksheet();
        $mergedCells = $this->worksheetType->getMergedCells();

        foreach ($mergedCells as $cells) {
            $sheet->mergeCells($cells)
                ->getStyle($cells)->getAlignment()->setHorizontal('center');
        }
    }
}
