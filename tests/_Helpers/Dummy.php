<?php

namespace Sleussink\ExcelWriter\Tests\_Helpers;

use DateTime;

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
