<?php

namespace Lysice\XlsWriter\Structure;

use Illuminate\Support\Facades\Log;
use Lysice\XlsWriter\Exceptions\Exception;
use Lysice\XlsWriter\Exceptions\ParamErrorException;

class Row {
    /**
     * @var \Vtiful\Kernel\Excel
     */
    protected $fileObject = null;
    /**
     * @var int
     */
    protected $rowIndex = 0;

    /**
     * @var array
     */
    protected $column = [];

    protected $formats = [];

    protected $urlText = '';

    protected $urlTooltip = '';

    /**
     * Row constructor.
     * @param $fileObject
     * @param array $column
     * @param $rowIndex
     * @param array $formats
     */
    public function __construct($fileObject, $column = [], $rowIndex = 1, $formats = [])
    {
        $this->fileObject = $fileObject;
        $this->column = $column;
        $this->rowIndex = $rowIndex;
        $this->formats = $formats;
    }

    /**
     * @return $this
     * @throws Exception
     */
    public function insert()
    {
        // urlText get
        if (isset($this->column['urlText'])) {
            $this->urlText = $this->column['urlText'];
            unset($this->column['urlText']);
        }
        // urlTooltip get
        if (isset($this->column['urlTooltip'])) {
            $this->urlTooltip = $this->column['urlTooltip'];
            unset($this->column['urlTooltip']);
        }

        foreach ($this->column as $cellIndex => $cellData) {
            if (empty($this->formats)) {
                TextCell::create($this, $cellIndex, $cellData, '', null)->insert();
            } else {
                $this->chooseCell($cellIndex, $this->formats, $cellData)->insert();
            }
        }

        return $this;
    }

    /**
     * choose cell by columnFormat
     * @param $cellIndex
     * @param $formats
     * @param $cellData
     * @return Cell
     * @throws Exception
     */
    public function chooseCell($cellIndex, $formats, $cellData) : Cell
    {
        if (!isset($formats[$cellIndex])) {
            throw new ParamErrorException('cell Format for index ' . $cellIndex . ' not defined.');
        }
        $format = isset($formats[$cellIndex]) ? $formats[$cellIndex] : null;
//        Log::error($cellData . '---' . $this->rowIndex . '--- ' . $cellIndex);
        $cell = null;
        switch ($format->cellType) {
            case CellConstants::CELL_TYPE_DATE:
                $dateFormat = isset($format->options['dateFormat']) ? $format->options['dateFormat'] : '';
                $cell = DateCell::create($this, $cellIndex, $cellData, $dateFormat, $format->getFormatResource($this->fileObject->getHandle()));
                break;
            case CellConstants::CELL_TYPE_FORMULA:
                $cell = FormulaCell::create($this, $cellIndex, $cellData, $format->getFormatResource($this->fileObject->getHandle()));
                break;
            case CellConstants::CELL_TYPE_IMAGE:
                $widthScale = isset($format->options['widthScale']) ? $format->options['widthScale'] : '';
                $heightScale = isset($format->options['heightScale']) ? $format->options['heightScale'] : '';

                if (!empty($widthScale) or !empty($heightScale)) {
                    if (empty($widthScale) or empty($heightScale)) {
                        throw new Exception('image scale must set at the same time');
                    }
                }
                $cell = ImageCell::create($this, $cellIndex, $cellData, $widthScale, $heightScale);
                break;
            case CellConstants::CELL_TYPE_URL:
                $cell = UrlCell::create($this, $cellIndex, $cellData, $this->urlText, $this->urlTooltip, $format->getFormatResource($this->fileObject->getHandle()));
                break;
            default:
                $cell = TextCell::create($this, $cellIndex, $cellData, '', $format->getFormatResource($this->fileObject->getHandle()));
                break;
        }

        return $cell;
    }

    private function cellParamsCheck()
    {

    }

    public function getRowIndex()
    {
        return $this->rowIndex;
    }

    public function getDelegate()
    {
        return $this->fileObject;
    }
}
