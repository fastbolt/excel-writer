<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\ColumnFormatter;

use Fastbolt\ExcelWriter\ColumnSetting;
use OutOfRangeException;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class CurrencyFormatter extends BaseFormatter
{
    public const CURRENCY_EUR = "EUR";

    public const CURRENCY_USD = "USD";

    private int $decimalLength;

    private string $currency;

    /**
     * @param ColumnSetting       $column
     * @param string{'EUR'|'USD'} $currency
     */
    public function __construct(ColumnSetting $column, string $currency)
    {
        $this->decimalLength = $column->getDecimalLength();
        $this->currency = $currency;
    }

    public function getAlignment(): array
    {
        return ['horizontal' => Alignment::HORIZONTAL_RIGHT];
    }

    public function getNumberFormat(): array
    {
        $formatCode = '0.' . str_repeat('0', $this->decimalLength);

        return ['formatCode' => $formatCode];
    }

    public function getFormatCode(): string
    {
        return match ($this->currency) {
            self::CURRENCY_EUR => NumberFormat::FORMAT_CURRENCY_EUR,
            self::CURRENCY_USD => NumberFormat::FORMAT_CURRENCY_USD,
            default => throw new OutOfRangeException("Currency  " . $this->currency . " is not supported"),
        };
    }
}
