<?php

namespace Sleussink\ExcelWriter\ColumnFormatter;

interface ColumnFormatter
{
    /** @return string[] */
    public function getAlignment(): array;

    /** @return string[]|null */
    public function getNumberFormat(): ?array;
}
