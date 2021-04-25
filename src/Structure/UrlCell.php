<?php

namespace Lysice\XlsWriter\Structure;

/**
 * Class UrlCell
 * @package Lysice\XlsWriter\Structure
 */
class UrlCell extends Cell {
    protected $tooltip = '';
    protected $text = '';
    protected $formatHandler;
    /**
     * UrlCell constructor.
     * @param $row
     * @param $cellIndex
     * @param $value
     * @param $text
     * @param $tooltip
     * @param $formatHandler
     */
    public function __construct($row, $cellIndex, $value, $text, $tooltip, $formatHandler)
    {
        parent::__construct($row, $cellIndex, $value);
        $this->tooltip = $tooltip;
        $this->text = $text;
        $this->formatHandler = $formatHandler;
    }

    public function insert()
    {
        if (empty($this->formatHandler)) {
            $this->row->getDelegate()->insertUrl($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->text, $this->tooltip, $this->formatHandler);
            return $this;
        }
        $this->row->getDelegate()->insertUrl($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->text, $this->tooltip, $this->formatHandler);
        return $this;
    }

    public static function create($row, $cellIndex, $value, $text, $tooltip, $formatHandler)
    {
        return new UrlCell($row, $cellIndex, $value, $text, $tooltip, $formatHandler);
    }
}
