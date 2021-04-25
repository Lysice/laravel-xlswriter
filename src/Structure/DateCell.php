<?php

namespace Lysice\XlsWriter\Structure;

use App\Http\Controllers\Auth\LoginController;
use Illuminate\Support\Facades\Log;
use Lysice\XlsWriter\Structure\Cell;

/**
 * Class DateCell
 * @package Structure
 */
class DateCell extends Cell{
    /**
     * the date for this cell
     * @var string
     */
    protected $dateFormat = '';

    protected $formatHandler;
    /**
     * DateCell constructor.
     * @param $row
     * @param $cellIndex
     * @param $value
     * @param $dateFormat
     * @param $formatHandler
     */
    public function __construct($row, $cellIndex, $value, $dateFormat, $formatHandler)
    {
        parent::__construct($row, $cellIndex, $value);
        $this->dateFormat = $dateFormat;
        $this->formatHandler = $formatHandler;
    }

    /**
     * insert date into cell
     * @return Cell|void
     */
    public function insert()
    {
        try {
            // todo: date cell set format error:502 bad gateway
            // see issue:https://github.com/viest/php-ext-xlswriter/issues/360
            // $this->row->getDelegate()->insertDate($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->dateFormat);
            $this->row->getDelegate()->insertDate($this->row->getRowIndex(), $this->cellIndex, $this->value, $this->dateFormat);
        } catch (\Exception $exception) {
            Log::error($exception->getMessage());
            Log::error($exception->getTraceAsString());
        }

        return $this;
    }

    public static function create($row, $cellIndex, $value, $dateFormat, $formatHandler)
    {
        return new DateCell($row, $cellIndex, $value, $dateFormat, $formatHandler);
    }
}
