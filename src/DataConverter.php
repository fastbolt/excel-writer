<?php

/**
 * Copyright © Fastbolt Schraubengroßhandels GmbH.
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Fastbolt\ExcelWriter;

use Exception;

class DataConverter
{
    /**
     * @param object[] $entities
     * @param ColumnSetting[] $cols
     * @throws Exception
     * @psalm-suppress MixedAssignment
     *
     * @return array
     */
    public function convertEntityToArray(array $entities, array $cols): array
    {
        //apply custom order of columns
        $getters = [];
        foreach ($cols as $col) {
            $getter = $col->getGetter();
            if (null === $getter || '' === $getter) {
                throw new Exception(
                    sprintf(
                        "All getters need to be set in the ColumnSettings when using entities. Missing getter for column '%s'",
                        $col->getHeader()
                    )
                );
            }

            $getters[] = $getter;
        }

        //call all getters
        $data = [];
        $counter = 0;
        foreach ($entities as $entity) {
            foreach ($getters as $getter) {
                //need to check for string, as is_callable lets php-methods pass (eg. 'getDate')
                if (is_callable($getter) && getType($getter) !== 'string') {
                    $data[$counter][] = $getter($entity);
                } else {
                    $value = $entity->$getter();
                    $data[$counter][] = $value;
                }
            }
            $counter++;
        }

        return $data;
    }

    /**
     * Calls the callable of the column for each row/item and replaces it with its return value
     * Callables of Entities should already have been replaced in convertEntityToArray()
     *
     * @param array           $data       indexed arrays
     * @param ColumnSetting[] $cols
     * @psalm-suppress MixedAssignment
     * @psalm-suppress MixedArrayAssignment
     * @psalm-suppress MixedArrayAccess
     *
     * @return array
     */
    public function resolveCallableGetters(array $data, array $cols): array
    {
        $colCounter = 0;
        foreach ($cols as $col) {
            $getter = $col->getGetter();
            if (is_callable($getter)) {
                $dataCounter = 0;
                foreach ($data as $item) {
                    $item[$colCounter] = $getter($item[$colCounter]);
                    $data[$dataCounter] = $item;
                    $dataCounter++;
                }
            }
            $colCounter++;
        }

        return $data;
    }
}
