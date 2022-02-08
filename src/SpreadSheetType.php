<?php

namespace Fastbolt\ExcelWriter;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SpreadSheetType
{
    private Spreadsheet $spreadsheet;

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

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->createSheet();
        $this->style = new TableStyle();
    }

    /**
     * @return Worksheet
     */
    public function getSheet(): Worksheet
    {
        return $this->spreadsheet->getActiveSheet();
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
     * @return SpreadSheetType
     */
    public function setSpreadsheet(Spreadsheet $spreadsheet): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setMaxColName(string $maxColName): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setMaxRowNumber(int $maxRowNumber): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setContentStartRow(int $contentStartRow): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setContent(array $content): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setStyle(TableStyle $style): SpreadSheetType
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
     * @return SpreadSheetType
     */
    public function setColumns(array $columns): SpreadSheetType
    {
        $this->columns = $columns;

        return $this;
    }

    /**
     * @param string $range
     * @return SpreadSheetType
     */
    public function setAutoFilterRange(string $range): SpreadSheetType
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
}
