<?php

namespace Lysice\XlsWriter\Structure;

/**
 * Class TextCell
 * @package Lysice\XlsWriter\Structure
 */
class TextCell extends Cell
{
    /**
     * the format for cell content
     * @var string
     */
    protected $format = '';

    protected $formatHandler;
    public function __construct($row, $cellIndex, $value, $format, $formatHandler)
    {
        parent::__construct($row, $cellIndex, $value);
        $this->format = $format;
        $this->formatHandler = $formatHandler;
    }

    /**
     * insert the data to cell with format
     * @return $this
     */
    public function insert()
    {
        if (empty($this->formatHandler) && empty($this->format)) {
            $this->row->getDelegate()->insertText($this->row->getRowIndex(), $this->cellIndex, $this->value);
            return $this;
        } else if (empty($this->formatHandler)) {
            $this->row->getDelegate()->insertText($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->format);
            return $this;
        } else if (empty($this->format)) {
            $this->row->getDelegate()->insertText($this->row->getRowIndex(), $this->cellIndex, $this->value, '', $this->formatHandler);
        }
        $this->row->getDelegate()->insertText($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->format, $this->formatHandler);
        return $this;
    }

    /**
     * get the text cell instance
     * @param $row
     * @param $cellIndex
     * @param $value
     * @param $format
     * @param $formatHandler
     * @return TextCell
     */
    public static function create($row, $cellIndex, $value, $format, $formatHandler) {
        return new TextCell($row, $cellIndex, $value, $format, $formatHandler);
    }
}
