<?php

namespace Fastbolt\ExcelWriter\Tests;

use Fastbolt\ExcelWriter\LetterProvider;
use PHPUnit\Framework\TestCase;

class LetterProviderTest extends TestCase
{
    public function testGetLetterForNumber(): void
    {
        $provider = new LetterProvider();

        self::assertEquals('A', $provider->getLetterForNumber(1));
        self::assertEquals('J', $provider->getLetterForNumber(10));
        self::assertEquals('Z', $provider->getLetterForNumber(26));
    }

    public function testGetNumberForLetter(): void
    {
        $provider = new LetterProvider();
        self::assertEquals(1, $provider->getNumberForLetter('A'));
        self::assertEquals(11, $provider->getNumberForLetter('K'));
        self::assertEquals(26, $provider->getNumberForLetter('Z'));
        self::assertFalse($provider->getNumberForLetter(''), 'no letter');
        self::assertFalse($provider->getNumberForLetter('AB'), 'too many letters');
    }
}
