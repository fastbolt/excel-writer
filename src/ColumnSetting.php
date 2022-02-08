<?php

namespace Fastbolt\ExcelWriter;

use Fastbolt\ExcelWriter\ColumnFormatter\ColumnFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\DateFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\FloatFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\IntegerFormatter;
use Fastbolt\ExcelWriter\ColumnFormatter\StringFormatter;

class ColumnSetting
{
    public const FORMAT_INTEGER = 'int';
    public const FORMAT_FLOAT   = 'float';
    public const FORMAT_STRING  = 'string';
    public const FORMAT_DATE    = 'datetime';

    private string $format;
    private string $name = '';    //excel-name for the column
    private string $header;       //heading of the column

    /**
     * @var callable|string name of the get method (like getId) or a callable taking the object an argument
     */
    private $getter;

    private int $decimalLength;

    /**
     * @param string          $header               column header
     * @param string          $format               format of the values
     * @param string|callable $getter               method name of the getter of the attribute or a callable
     * @param int             $decimalLength        only for float columns: how many decimals are displayed
     */
    public function __construct(
        string $header,
        string $format = self::FORMAT_STRING,
        $getter = '',
        int $decimalLength = 2
    ) {
        $this->header        = $header;
        $this->format        = $format;
        $this->getter        = $getter;
        $this->decimalLength = $decimalLength;
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

    /**
     * @return string|callable
     */
    public function getGetter()
    {
        return $this->getter;
    }

    /**
     * @param string|callable $getter
     * @return $this
     */
    public function setGetter($getter): self
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

    public function setDecimalLength(int $count): ColumnSetting
    {
        $this->decimalLength = $count;

        return $this;
    }
}
