<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Fastbolt\ExcelWriter\ColumnSetting;

class FloatFormatter extends BaseFormatter
{
    private int $decimalLength;

    public function __construct(ColumnSetting $column)
    {
        $this->decimalLength = $column->getDecimalLength();
    }

    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_RIGHT];
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
