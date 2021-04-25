<?php

namespace Lysice\XlsWriter\Structure;

class Cell {
    /**
     * @var null
     */
    protected $value = null;
    /**
     * @var Row
     */
    protected $row;

    /**
     * the index in column where you insert the data to.
     * @var int
     */
    protected $cellIndex = 0;

    /**
     * @var Resource
     */
    protected $formatHandler;

    /**
     * Cell constructor.
     * @param $row
     * @param $cellIndex
     * @param $value
     */
    public function __construct($row, $cellIndex, $value)
    {
        $this->row = $row;
        $this->cellIndex = $cellIndex;
        $this->value = $value;
    }

    public function insert()
    {
        return $this;
    }
}
