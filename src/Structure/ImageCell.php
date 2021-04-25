<?php

namespace Lysice\XlsWriter\Structure;

class ImageCell extends Cell {
    /**
     * @var float
     */
    protected $widthScale = 1.0;
    /**
     * @var float
     */
    protected $heightScale = 1.0;

    public function __construct($row, $cellIndex, $value, $widthScale, $heightScale)
    {
        parent::__construct($row, $cellIndex, $value);
        $this->widthScale = $widthScale;
        $this->heightScale = $heightScale;
    }

    /**
     * @return $this|Cell
     */
    public function insert()
    {
        if (!empty($this->widthScale) && !empty($this->heightScale)) {
            $this->row->getDelegate()->insertImage($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->widthScale, $this->heightScale);
        } else {
            $this->row->getDelegate()->insertImage($this->row->getRowIndex(), $this->cellIndex, $this->value);
        }

        return $this;
    }

    public static function create($row, $cellIndex, $value, $widthScale, $heightScale)
    {
        return new ImageCell($row, $cellIndex, $value, $widthScale, $heightScale);
    }
}
