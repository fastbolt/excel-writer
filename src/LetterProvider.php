<?php

namespace Fastbolt\ExcelWriter;

class LetterProvider
{
    /**
     * @var string[]
     */
    private array $alphabet;

    public function __construct()
    {
        $this->alphabet = range('A', 'Z');
    }

    /**
     * @param int $number
     *
     * @return string
     */
    public function getLetterForNumber(int $number): string
    {
        return $this->alphabet[$number - 1];
    }

    /**
     * @param string $letter
     *
     * @return false|int
     */
    public function getNumberForLetter(string $letter)
    {
        if (strlen($letter) !== 1) {
            return false;
        }

        return (int) array_search($letter, $this->alphabet, false) + 1;
    }
}
