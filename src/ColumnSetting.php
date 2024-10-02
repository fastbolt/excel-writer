<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter;

use Fastbolt\ExcelWriter\ColumnFormatter\ColumnFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\CurrencyFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\DateFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\FloatFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\IntegerFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\PercentageFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;
use PhpOffice\PhpSpreadsheet\Calculation\DateTimeExcel\Current;

class ColumnSetting
{
    public const FORMAT_INTEGER    = 'int';
    public const FORMAT_FLOAT      = 'float';
    public const FORMAT_STRING     = 'string';
    public const FORMAT_DATE       = 'datetime';
    public const FOMRAT_PERCENTAGE = 'percentage';
    public const FORMAT_CURRENCY   = 'currency';

    private string $format;
    private string $name = '';    //excel-name for the column
    private string $header;       //heading of the column

    private int $decimalLength;

    private ?array $headerStyle;
    private ?array $dataStyle;

    /** @var string|callable|null name of the get method (like getId) or a callable taking the object an argument */
    private $getter;


    /**
     * @param string          $header               column header
     * @param string          $format               format of the values
     * @param string|callable|null $getter          method name of the getter of the attribute or a callable
     * @param int             $decimalLength        only for float columns: how many decimals are displayed
     */
    public function __construct(
        string $header,
        string $format = self::FORMAT_STRING,
        string|callable|null $getter = '',
        int $decimalLength = 2,
        ?array $headerStyle = null,
        ?array $dataStyle = null
    ) {
        $this->header           = $header;
        $this->format           = $format;
        $this->getter           = $getter;
        $this->decimalLength    = $decimalLength;
        $this->headerStyle      = $headerStyle;
        $this->dataStyle        = $dataStyle;
    }

    /**
     * @return string
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * names is set automatically when headers are set
     *
     * @param string $name
     *
     * @return ColumnSetting
     */
    public function setName(string $name): ColumnSetting
    {
        $this->name = $name;

        return $this;
    }

    /**
     * @return ColumnFormatter
     */
    public function getFormatter(): ColumnFormatter
    {
        switch ($this->format) {
            case self::FORMAT_STRING:
                return new StringFormatter();
            case self::FORMAT_DATE:
                return new DateFormatter();
            case self::FORMAT_INTEGER:
                return new IntegerFormatter();
            case self::FORMAT_FLOAT:
                return new FloatFormatter($this);
            case self::FOMRAT_PERCENTAGE:
                return new PercentageFormatter($this);
            case self::FORMAT_CURRENCY:
                return new CurrencyFormatter($this);
        }

        return new StringFormatter();
    }

    /**
     * @param string $format
     *
     * @return ColumnSetting
     */
    public function setFormat(string $format): ColumnSetting
    {
        $this->format = $format;

        return $this;
    }

    public function getGetter(): string|callable|null
    {
        return $this->getter;
    }

    public function setGetter(string|callable|null $getter): self
    {
        $this->getter = $getter;

        return $this;
    }

    /**
     * @return string
     */
    public function getHeader(): string
    {
        return $this->header;
    }

    /**
     * @param string $header
     * @return ColumnSetting
     */
    public function setHeader(string $header): ColumnSetting
    {
        $this->header = $header;

        return $this;
    }

    /**
     * returns how many decimal places are displayed if the column format is float
     */
    public function getDecimalLength(): int
    {
        return $this->decimalLength;
    }

    /**
     * @param int $count
     * @return $this
     */
    public function setDecimalLength(int $count): ColumnSetting
    {
        $this->decimalLength = $count;

        return $this;
    }

    /**
     * @param array|null $style
     * @return ColumnSetting
     */
    public function setHeaderStyle(?array $style): ColumnSetting
    {
        $this->headerStyle = $style;

        return $this;
    }

    /**
     * @return array|null
     */
    public function getHeaderStyle(): ?array
    {
        return $this->headerStyle;
    }


    /**
     * @param array|null $style
     * @return ColumnSetting
     */
    public function setDataStyle(?array $style): ColumnSetting
    {
        $this->dataStyle = $style;

        return $this;
    }

    /**
     * @return array|null
     */
    public function getDataStyle(): ?array
    {
        return $this->dataStyle;
    }
}
