<?php

namespace Lysice\XlsWriter\Structure;

class FormulaCell extends Cell{
    protected $formatHandler;
    public function __construct($row, $cellIndex, $value, $formatHandler)
    {
        parent::__construct($row, $cellIndex, $value);
        $this->formatHandler = $formatHandler;
    }

    /**
     * @return $this|Cell
     */
    public function insert() {
        $this->row->getDelegate()->insertFormula($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->formatHandler);
        return $this;
    }

    public static function create($row, $cellIndex, $value, $formatHandler)
    {
        return new FormulaCell($row, $cellIndex, $value, $formatHandler);
    }
}
