<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter;

class TableStyle
{
    private int $headerRowHeight = 1;

    private array $headerStyle = [];

    private array $dataRowStyle = [];

    public function setHeaderRowHeight(int $headerRows): TableStyle
    {
        $this->headerRowHeight = $headerRows;

        return $this;
    }

    /**
     * @return int
     */
    public function getHeaderRowHeight(): int
    {
        return $this->headerRowHeight;
    }

    public function getHeaderStyle(): array
    {
        return $this->headerStyle;
    }

    public function setHeaderStyle(array $headerStyle): TableStyle
    {
        $this->headerStyle = $headerStyle;

        return $this;
    }

    public function getDataRowStyle(): array
    {
        return $this->dataRowStyle;
    }

    public function setDataRowStyle(array $rowStyle): TableStyle
    {
        $this->dataRowStyle = $rowStyle;

        return $this;
    }
}
