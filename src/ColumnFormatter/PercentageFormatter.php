<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use Fastbolt\ExcelWriter\ColumnSetting;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class PercentageFormatter extends BaseFormatter
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

    public function getNumberFormat(): array
    {
        $formatCode = '0.' . str_repeat('0', $this->decimalLength) . '%';

        return ['formatCode' => $formatCode];
    }
}
