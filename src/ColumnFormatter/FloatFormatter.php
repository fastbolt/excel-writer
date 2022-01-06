<?php

namespace Sleussink\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Sleussink\ExcelWriter\ColumnSetting;

class FloatFormatter extends BaseFormatter
{
    private int $decimalLength;

    public function __construct(ColumnSetting $column)
    {
        $this->decimalLength = $column->getDecimalLength();
    }

    public function getAlignment(): array
    {
        return ['vertical' => Alignment::HORIZONTAL_RIGHT];
    }

    /**
     * @return string[]
     */
    public function getNumberFormat(): array
    {
        $formatCode = '0.'.str_repeat('0', $this->decimalLength);

        return ['formatCode' => $formatCode];
    }
}
