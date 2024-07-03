<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

class WorksheetType
{
    private ?Worksheet $worksheet = null;

    private string $title = 'Worksheet';

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

    /**
     * @var string[]
     */
    private array $mergedCells = [];

    private Spreadsheet $spreadsheet;

    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
        $this->style = new TableStyle();
    }

    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    public function setSpreadsheet(Spreadsheet $spreadsheet): WorksheetType
    {
        $this->spreadsheet = $spreadsheet;

        return $this;
    }

    public function getMaxColName(): string
    {
        return $this->maxColName;
    }

    public function setMaxColName(string $maxColName): WorksheetType
    {
        $this->maxColName = $maxColName;

        return $this;
    }

    public function getMaxRowNumber(): int
    {
        return $this->maxRowNumber;
    }

    public function setMaxRowNumber(int $maxRowNumber): WorksheetType
    {
        $this->maxRowNumber = $maxRowNumber;

        return $this;
    }

    public function getContentStartRow(): int
    {
        return $this->contentStartRow;
    }

    public function setContentStartRow(int $contentStartRow): WorksheetType
    {
        $this->contentStartRow = $contentStartRow;

        return $this;
    }

    public function getContent(): array
    {
        return $this->content;
    }

    public function setContent(array $content): WorksheetType
    {
        $this->content = $content;

        return $this;
    }

    public function getStyle(): TableStyle
    {
        return $this->style;
    }

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
     */
    public function setColumns(array $columns): WorksheetType
    {
        $this->columns = $columns;

        return $this;
    }

    public function setAutoFilterRange(string $range): WorksheetType
    {
        if (!Coordinate::coordinateIsRange($range)) {
            throw new \InvalidArgumentException('Method setAutoFilterRange() must be passed a range');
        }

        $this->autoFilterRange = $range;

        return $this;
    }

    public function getAutoFilterRange(): string
    {
        return $this->autoFilterRange;
    }

    /**
     * @param string[] $mergedCells
     */
    public function setMergedCells(array $mergedCells): WorksheetType
    {
        $this->mergedCells = $mergedCells;

        return $this;
    }

    /**
     * @param string[] $mergedCells
     */
    public function addMergedCells(array $mergedCells): WorksheetType
    {
        foreach ($mergedCells as $cells) {
            $this->mergedCells[] = $cells;
        }

        return $this;
    }

    /**
     * @return string[]
     */
    public function getMergedCells(): array
    {
        return $this->mergedCells;
    }

    public function getWorksheet(): Worksheet
    {
        return $this->worksheet ?: throw new RuntimeException('Worksheet not set');
    }

    public function setWorksheet(Worksheet $worksheet): void
    {
        $this->worksheet = $worksheet;
    }

    public function getTitle(): string
    {
        return $this->title;
    }

    public function setTitle(string $title): void
    {
        $this->title = $title;
    }
}
