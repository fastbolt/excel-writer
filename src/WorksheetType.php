<?php

namespace Fastbolt\ExcelWriter;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class WorksheetType
{
    private Worksheet $worksheet;

    private $title = 'Worksheet';

    private string $maxColName = '';

    private int $maxRowNumber = 0;

    private int $contentStartRow = 2;

    /**
     * @var ColumnSetting[]
     */
    private array $columns = [];

    private array $content = [];

    private TableStyle $style;

    private string $autoFilterRange = '';

    private array $mergedCells = [];

    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->setSpreadsheet($spreadsheet);
        $this->style = new TableStyle();
    }

    /**
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * @param Spreadsheet $spreadsheet
     *
     * @return WorksheetType
     */
    public function setSpreadsheet(Spreadsheet $spreadsheet): WorksheetType
    {
        $this->spreadsheet = $spreadsheet;

        return $this;
    }


    /**
     * @return string
     */
    public function getMaxColName(): string
    {
        return $this->maxColName;
    }

    /**
     * @param string $maxColName
     *
     * @return WorksheetType
     */
    public function setMaxColName(string $maxColName): WorksheetType
    {
        $this->maxColName = $maxColName;

        return $this;
    }

    /**
     * @return int
     */
    public function getMaxRowNumber(): int
    {
        return $this->maxRowNumber;
    }

    /**
     * @param int $maxRowNumber
     *
     * @return WorksheetType
     */
    public function setMaxRowNumber(int $maxRowNumber): WorksheetType
    {
        $this->maxRowNumber = $maxRowNumber;

        return $this;
    }

    /**
     * @return int
     */
    public function getContentStartRow(): int
    {
        return $this->contentStartRow;
    }

    /**
     * @param int $contentStartRow
     *
     * @return WorksheetType
     */
    public function setContentStartRow(int $contentStartRow): WorksheetType
    {
        $this->contentStartRow = $contentStartRow;

        return $this;
    }

    /**
     * @return array
     */
    public function getContent(): array
    {
        return $this->content;
    }

    /**
     * @param array $content
     *
     * @return WorksheetType
     */
    public function setContent(array $content): WorksheetType
    {
        $this->content = $content;

        return $this;
    }

    /**
     * @return TableStyle
     */
    public function getStyle(): TableStyle
    {
        return $this->style;
    }

    /**
     * @param TableStyle $style
     *
     * @return WorksheetType
     */
    public function setStyle(TableStyle $style): WorksheetType
    {
        $this->style = $style;

        return $this;
    }

    /**
     * @return ColumnSetting[]
     */
    public function getColumns(): array
    {
        return $this->columns;
    }

    /**
     * @param ColumnSetting[] $columns
     *
     * @return WorksheetType
     */
    public function setColumns(array $columns): WorksheetType
    {
        $this->columns = $columns;

        return $this;
    }

    /**
     * @param string $range
     * @return WorksheetType
     */
    public function setAutoFilterRange(string $range): WorksheetType
    {
        if (!Coordinate::coordinateIsRange($range)) {
            throw new \InvalidArgumentException('Method setAutoFilterRange() must be passed a range');
        }

        $this->autoFilterRange = $range;

        return $this;
    }

    /**
     * @return string
     */
    public function getAutoFilterRange(): string
    {
        return $this->autoFilterRange;
    }

    /**
     * @param array $mergedCells
     * @return WorksheetType
     */
    public function setMergedCells(array $mergedCells): WorksheetType
    {
        $this->mergedCells = $mergedCells;

        return $this;
    }

    /**
     * @param string[] $mergedCells
     * @return WorksheetType
     */
    public function addMergedCells(array $mergedCells): WorksheetType
    {
        foreach ($mergedCells as $cells) {
            $this->mergedCells[] = $cells;
        }

        return $this;
    }

    /**
     * @return array
     */
    public function getMergedCells(): array
    {
        return $this->mergedCells;
    }

    /**
     * @return Worksheet
     */
    public function getWorksheet(): Worksheet
    {
        return $this->worksheet;
    }

    /**
     * @param Worksheet $worksheet
     */
    public function setWorksheet(Worksheet $worksheet): void
    {
        $this->worksheet = $worksheet;
    }

    /**
     * @return string
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * @param string $title
     */
    public function setTitle(string $title): void
    {
        $this->title = $title;
    }
}
