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

    /**
     * @param int $headerRows
     *
     * @return TableStyle
     */
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

    /**
     * @return array[]
     */
    public function getHeaderStyle(): array
    {
        return $this->headerStyle;
    }

    /**
     * @param array[] $headerStyle
     *
     * @return TableStyle
     */
    public function setHeaderStyle(array $headerStyle): TableStyle
    {
        $this->headerStyle = $headerStyle;

        return $this;
    }

    /**
     * @return array[]
     */
    public function getDataRowStyle(): array
    {
        return $this->dataRowStyle;
    }

    /**
     * @param array[] $rowStyle
     *
     * @return TableStyle
     */
    public function setDataRowStyle(array $rowStyle): TableStyle
    {
        $this->dataRowStyle = $rowStyle;

        return $this;
    }
}
