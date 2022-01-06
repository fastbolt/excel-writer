<?php

namespace Sleussink\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;

class DateFormatter extends BaseFormatter
{
    public function getAlignment(): array
    {
        return ['vertical' => Alignment::HORIZONTAL_LEFT];
    }
}
