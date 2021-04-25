<?php

namespace Lysice\XlsWriter\Supports\Format;

use Lysice\XlsWriter\Exceptions\ParamErrorException;
use Lysice\XlsWriter\Structure\CellConstants;

/**
 * Class ColumnFormat
 * @package Lysice\XlsWriter\Supports\Format
 */
class ColumnFormat extends DefaultFormat {
    public $cellTypes = [
        CellConstants::CELL_TYPE_TEXT,
        CellConstants::CELL_TYPE_DATE,
        CellConstants::CELL_TYPE_FORMULA,
        CellConstants::CELL_TYPE_IMAGE,
        CellConstants::CELL_TYPE_URL,
    ];

    /**
     * @var int|null
     */
    public $cellType = CellConstants::CELL_TYPE_TEXT;

    /**
     * @var array | null
     */
    public $options = [];

    /**
     * set the cell type in this column
     * @param $type
     * @return $this
     */
    public function setCellType($type)
    {
        if (!in_array($type, $this->cellTypes)) {
            throw new ParamErrorException("cell type does't support!");
        }

        $this->cellType = $type;
        return $this;
    }

    /**
     *
     * @param array $options
     * @return $this
     */
    public function setOptions(array $options = [])
    {
        $this->options = $options;
        return $this;
    }
}
