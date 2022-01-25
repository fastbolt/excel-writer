<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;

class BaseFormatter implements ColumnFormatter
{
    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_LEFT];
    }

    public function getNumberFormat(): ?array
    {
        return null;
    }
}
