<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\ColumnFormatter;

interface ColumnFormatter
{
    /** @return string[] */
    public function getAlignment(): array;

    /** @return string[]|null */
    public function getNumberFormat(): ?array;
}
