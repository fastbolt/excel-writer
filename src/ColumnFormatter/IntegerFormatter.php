<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class IntegerFormatter extends BaseFormatter
{
    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_RIGHT];
    }

    public function getNumberFormat(): ?array
    {
        return ['formatCode' => NumberFormat::FORMAT_NUMBER];
    }
}
