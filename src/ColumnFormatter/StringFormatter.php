<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;

class StringFormatter extends BaseFormatter
{
    public function getAlignment(): array
    {
        return ['vertical' => Alignment::HORIZONTAL_LEFT];
    }
}
