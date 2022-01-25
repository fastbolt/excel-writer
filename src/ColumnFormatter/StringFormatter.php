<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class StringFormatter extends BaseFormatter
{
    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_LEFT];
    }

    public function getNumberFormat(): array
    {
        return ['formatCode' => NumberFormat::FORMAT_TEXT];
    }
}
