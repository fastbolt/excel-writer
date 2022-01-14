<?php

namespace Fastbolt\ExcelWriter\ColumnFormatter;

interface ColumnFormatter
{
    /** @return string[] */
    public function getAlignment(): array;

    /** @return string[]|null */
    public function getNumberFormat(): ?array;
}
