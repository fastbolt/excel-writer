<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter\Tests;

use DateTime;
use Fastbolt\ExcelWriter\ColumnSetting;
use Fastbolt\ExcelWriter\DataConverter;
use Fastbolt\ExcelWriter\Tests\_Helpers\Dummy;
use PHPUnit\Framework\TestCase;

/**
 * @covers \Fastbolt\ExcelWriter\DataConverter
 */
class DataConverterTest extends TestCase
{
    private Dummy $object;

    public function setUp(): void
    {
        $this->object = new Dummy();
    }

    public function testConvertEntityToArrayNotAllGettersSetError(): void
    {
        $cols = [
          new ColumnSetting('name', ColumnSetting::FORMAT_INTEGER, 'getName'),
          new ColumnSetting('value'),
          new ColumnSetting('date')
        ];

        $converter = new DataConverter();

        $this->expectExceptionMessage(
            'All getters need to be set in the ColumnSettings when using entities. Missing getter for column \'value\''
        );
        $converter->convertEntityToArray([$this->object], $cols);
    }

    public function testConvertEntityToArrayAllGettersSet(): void
    {
        $cols = [
            new ColumnSetting('value', ColumnSetting::FORMAT_INTEGER, 'getValue'),
            new ColumnSetting('date', ColumnSetting::FORMAT_DATE, 'getDate'),
            new ColumnSetting('name', ColumnSetting::FORMAT_STRING, 'getName'),
            new ColumnSetting('bool', ColumnSetting::FORMAT_INTEGER, 'isTrue')
        ];

        $converter = new DataConverter();
        $result = $converter->convertEntityToArray([$this->object], $cols);
        $item = $result[0];

        self::assertEquals(100, $item[0], 'value');
        self::assertInstanceOf(DateTime::class, $item[1], 'date');
        self::assertEquals('name', $item[2], 'name');
        self::assertEquals('1', $item[3], 'bool');
    }

    public function testConvertEntityToArrayWithCallableGetter(): void
    {
        $cols = [
            new ColumnSetting('value', ColumnSetting::FORMAT_INTEGER, static function () {
                return 'foo';
            })
        ];
        $converter = new DataConverter();
        $return = $converter->convertEntityToArray([new Dummy()], $cols);

        self::assertEquals('foo', $return[0][0]);
    }

    public function testResolveCallableGetters(): void
    {
        $data = [
            [
                'foo',
                new Dummy()
            ]
        ];

        $cols = [
            new ColumnSetting('notCallable', ''),
            new ColumnSetting('callable', '', static function ($dummy) {
                return $dummy->getValue();
            })
        ];

        $converter = new DataConverter();
        $result = $converter->resolveCallableGetters($data, $cols);

        self::assertEquals('foo', $result[0][0], 'not callable');
        self::assertEquals(100, $result[0][1], 'callable');
    }
}
