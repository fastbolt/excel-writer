<?php

namespace Fastbolt\ExcelWriter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SplFileInfo;

class ExcelGenerator
{
    //TODO set the apply... and save functions to private, adapt tests
    private LetterProvider  $letterProvider;
    private SpreadSheetType $spreadsheetType;
    private DataConverter $converter;

    public function __construct(
        SpreadSheetType $spreadsheetType,
        ?DataConverter $converter = null,
        ?LetterProvider $letterProvider = null
    ) {
        $this->spreadsheetType  = $spreadsheetType;
        $this->converter        = $converter ?? new DataConverter();
        $this->letterProvider   = $letterProvider ?? new LetterProvider();
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

        $colName = $this->letterProvider->getLetterForNumber(count($columns));
        $this->spreadsheetType->setMaxColName($colName);

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
            throw new \ArgumentCountError('At least one column must be set.');
        }

        $this->applyColumnHeaders($this->spreadsheetType->getColumns());
        $this->applyColumnFormat($this->spreadsheetType->getColumns());
        $this->applyHeaderStyle($this->spreadsheetType->getStyle());

        if ($this->spreadsheetType->getContent()) {
            $this->applyContent($this->spreadsheetType->getContent());
        }

        $this->applyStyle($this->spreadsheetType->getStyle());

        //auto filter
        $sheet = $this->spreadsheetType->getSheet();
        $sheet->setAutoFilter($sheet->calculateWorksheetDimension());

        //auto size
        $maxCol = $this->spreadsheetType->getMaxColName();
        foreach (range('A', $maxCol) as $column) {
            $sheet->getColumnDimension($column)
                ->setAutoSize(true);
        }

        return $this->saveFile($url);
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
            $url = sys_get_temp_dir() . '/spreadsheet';
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
                $colName     = $cols[$counter]->getName();
                $coordinates = $colName . $currentRow;
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
    public function applyStyle(TableStyle $style): ExcelGenerator
    {
        //this should probably be made coordinate / item specific

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
     * @param ColumnSetting[] $columns
     *
     * @return ExcelGenerator
     */
    public function applyColumnHeaders(array $columns): ExcelGenerator
    {
        $sheet       = $this->spreadsheetType->getSheet();
        $headerCount = count($columns);
        $last        = $this->letterProvider->getLetterForNumber($headerCount);
        $letters     = range('A', $last);

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
            $cols = range('A', $this->spreadsheetType->getMaxColName());
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
                'borders'   => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_MEDIUM,
                        'color'       => ['argb' => 'FF366092'],
                    ],
                ],
            ])
            ->applyFromArray($style->getHeaderStyle());

        return $this;
    }
}
