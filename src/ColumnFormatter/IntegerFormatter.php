<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class IntegerFormatter implements ColumnFormatter
{
    public function getAlignment(): array
    {
        return ['vertical' => Alignment::HORIZONTAL_RIGHT];
    }

    public function getNumberFormat(): ?array
    {
        return ['formatCode' => NumberFormat::FORMAT_NUMBER];
    }
}
