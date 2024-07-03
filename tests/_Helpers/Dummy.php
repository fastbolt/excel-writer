<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests\_Helpers;

use DateTime;

/**
 * @coversNothing
 */
class Dummy
{
    public function getName(): string
    {
        return 'name';
    }

    public function getValue(): int
    {
        return 100;
    }

    public function getDate(): DateTime
    {
        return new DateTime();
    }

    public function isTrue(): bool
    {
        return true;
    }

    public function getObject(): self
    {
        return $this;
    }
}
