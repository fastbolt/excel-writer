<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

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

    public function getFormatCode(): ?string
    {
        return null;
    }
}
