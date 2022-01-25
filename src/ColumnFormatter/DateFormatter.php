<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;

class DateFormatter extends BaseFormatter
{
    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_LEFT];
    }
}
