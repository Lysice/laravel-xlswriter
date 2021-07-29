<?php

namespace Lysice\XlsWriter;

use Lysice\XlsWriter\Exceptions\AutoFilterModeNotSupportException;
use Lysice\XlsWriter\Exceptions\ConfigNotFoundException;
use Lysice\XlsWriter\Exceptions\Exception;
use Lysice\XlsWriter\Exceptions\ExportMustImplementException;
use Lysice\XlsWriter\Exceptions\ParamErrorException;
use Lysice\XlsWriter\FileSystem\ExportedFile;
use Lysice\XlsWriter\Interfaces\FromArray;
use Lysice\XlsWriter\Interfaces\FromCollection;
use Lysice\XlsWriter\Interfaces\FromQuery;
use Lysice\XlsWriter\Interfaces\WithCharts;
use Lysice\XlsWriter\Interfaces\WithDefaultFormat;
use Lysice\XlsWriter\Interfaces\WithFilter;
use Lysice\XlsWriter\Interfaces\WithGridLine;
use Lysice\XlsWriter\Interfaces\WithColumnFormat;
use Lysice\XlsWriter\Interfaces\WithZoom;
use Lysice\XlsWriter\Structure\Row;
use Lysice\XlsWriter\Supports\Chart\ChartFactory;
use Lysice\XlsWriter\Supports\Format\DefaultFormat;
use Lysice\XlsWriter\Supports\Format\Format;
use Lysice\XlsWriter\Traits\CommonTrait;
use Lysice\XlsWriter\Traits\ExcelExportTrait;

/**
 * Class XlsWriter
 * @package Lysice\XlsWriter
 */
class XlsWriter {
    use ExcelExportTrait, CommonTrait;

    /**
     * @var array
     */
    protected $config;

    /**
     * @var \Vtiful\Kernel\Excel
     */
    protected $excel;

    /**
     * @var array
     */
    protected $header;

    /**
     * @var array
     */
    protected $data;

    /**
     * @var string
     */
    protected $filePath;

    /**
     * @var string
     */
    protected $protection = '';

    /**
     * @var string
     */
    protected $path = '';

    /**
     * @var string
     */
    protected $filter;

    /**
     * @var Format
     */
    protected $defaultHeaderFormat;

    /**
     * @var Format
     */
    protected $defaultContentFormat;

    /**
     * @var integer
     */
    protected $printMode = \Lysice\XlsWriter\Excel::MODE_PRINT_PORTRAIT;

    protected $charts;

    protected $gridLine;

    protected $zoom;

    /**
     * @var array | null
     */
    protected $rowFormats;
    /**
     * XlsWriter constructor.
     * @throws ConfigNotFoundException
     */
    public function __construct()
    {
        $this->config = config('xlswriter');
        if (empty($this->config)) {
            throw new ConfigNotFoundException('laravel-xlswriter config file is not published.please publish first');
        }
    }

    /**
     * @param $export
     * @param $writerType
     * @param $fileName
     * @return ExportedFile
     * @throws Exception
     */
    public function export($export, $writerType, $fileName)
    {
        $fileName = $fileName . '.'. $writerType;
        $filePath = $this->handleExport($this->parseExport($export), $fileName);
        return new ExportedFile($filePath);
    }

    /**
     * @param $export
     * @return array|\Illuminate\Support\Collection
     * @throws Exception
     */
    protected function parseExport($export)
    {
        $this->header = $export->headers();
        $this->setProtection($export->protection ?? '');
        $this->setPrintMode($export->printMode ?? \Lysice\XlsWriter\Excel::MODE_PRINT_PORTRAIT);
        if ($export instanceof WithCharts) {
            $this->setCharts($export);
        }
        if ($export instanceof WithFilter) {
            $this->setFilter($export);
        }
        if ($export instanceof WithDefaultFormat) {
            $this->setDefaultFormat($export);
        }
        if ($export instanceof WithGridLine) {
            $this->setGridLine($export);
        }

        if ($export instanceof WithZoom) {
            $this->setZoom($export);
        }
        if ($export instanceof WithColumnFormat) {
            $this->setRowFormats($export);
        }
        if ($export instanceof FromArray) {
            return $export->array();
        } else if ($export instanceof FromCollection) {
            $result = [];
            foreach ($export->collection()->toArray() as $collections) {
                $result[] = array_merge(array_values((array)$collections), $result);
            }
            return $result;
        } else if ($export instanceof FromQuery) {
            return [];
        }
        else {
            throw new ExportMustImplementException();
        }
    }

    protected function setGridLine(WithGridLine $export)
    {
        $this->gridLine = $export->gridLine();
        return $this;
    }

    protected function setRowFormats(WithColumnFormat $export)
    {
        $this->rowFormats = $export->columnFormats();
        return $this;
    }

    protected function setZoom(WithZoom $export) {
        $this->zoom = $export->zoom();
        return $this;
    }

    protected function setFilter(WithFilter $export)
    {
        $this->filter = $export->filter();
        return $this;
    }

    protected function autoFilter()
    {
        if ($this->config['mode'] !=  \Lysice\XlsWriter\Excel::MODE_NORMAL) {
            throw new AutoFilterModeNotSupportException('In const memory mode, can not use autoFilter');
        }

        $this->excel->autoFilter($this->filter);
        return $this;
    }

    protected function setCharts(WithCharts $export)
    {
        $this->charts = $export->charts();
        return $this;
    }

    /**
     * export excel function
     * @param string $fileName
     * @param array $header
     * @param array $data
     * @return XlsWriter|mixed
     * @throws Exception
     */
    public function handleExport($data = [], $fileName = '')
    {
        if (empty($data)) {
            throw new Exception('the export document data is not set!');
        }

        try {
            return $this->getExcelInstance($fileName)
                ->when($this->defaultHeaderFormat, function ($xlsWriter, $defaultHeaderFormat) {
                    $xlsWriter->defaultFormat($defaultHeaderFormat);
                    return $xlsWriter;
                })
                ->when($this->header, function ($xlsWriter){
                    $xlsWriter->setHeader($this->header);
                    return $xlsWriter;
                })
                ->when($this->defaultContentFormat, function ($xlsWriter, $defaultContentFormat) {
                    $xlsWriter->defaultFormat($defaultContentFormat);
                    return $xlsWriter;
                })
                ->when(isset($this->gridLine), function ($xlsWriter) {
                    return $xlsWriter->gridLine();
                })
                ->when($this->zoom, function ($xlsWriter, $zoom) {
                    return $xlsWriter->zoom($zoom);
                })
                ->when($this->config['export_mode'], function ($xlsWriter, $mode) use ($data) {
                    if ($mode == Excel::EXPORT_MODE_DATA) {
                        return $xlsWriter->modeData($data);
                    }

                    return $xlsWriter->modeRow($data);
                })

                ->when($this->filter, function ($xlsWriter, $filters) {
                    return $xlsWriter->autoFilter();
                })
                ->when($this->protection, function ($xlsWriter, $protection) {
                    return $xlsWriter->protection($protection);
                })->when($this->printMode, function ($xlsWriter, $printMode) {
                    if ($printMode == 2) {
                        return $xlsWriter->setPortrait();
                    }
                    return $xlsWriter->setLandscape();
                })
                ->when($this->charts, function ($xlsWriter, $charts) {
                    return $xlsWriter->insertCharts($charts);
                })->output();
        } catch (\Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /**
     * get the excel instance by mode
     * @param string $fileName
     * @return $this
     * @throws Exception
     */
    private function getExcelInstance($fileName = '')
    {
        if (isset($this->config['mode']) && $this->config['mode'] == \Lysice\XlsWriter\Excel::MODE_NORMAL) {
            return $this->getExcelNormalInstance($fileName);
        }

        return $this->getExcelMemoryInstance($fileName);
    }

    /**
     * get the excel instance which runs in memory
     * @param string $fileName
     * @return $this
     * @throws Exception
     */
    private function getExcelMemoryInstance($fileName = '')
    {
        $this->excel()->memory($fileName);
        return $this;
    }

    /**
     * add charts to excel
     * @param array $chartConfigs
     * @return $this
     * @throws Exceptions\ChartParamErrorException
     */
    protected function insertCharts(array $chartConfigs = [])
    {
        $charts = ChartFactory::makeCharts($this->excel->getHandle(), $chartConfigs);
        foreach ($charts as $chart) {
            $this->excel->insertChart($chart->row, $chart->column, $chart->toResource());
        }

        return $this;
    }

    private function zoom($scale = 0) {
        if (!is_numeric($scale)) {
            throw new ParamErrorException('zoom scale param must be a number');
        }
        if ($scale < 10 or $scale > 400) {
            throw new ParamErrorException('zoom scale param error!');
        }
        $this->excel->zoom($scale);
        return $this;
    }

    private function gridLine()
    {
        $gridLine = $this->gridLine;
        if (!is_numeric($gridLine)) {
            throw new ParamErrorException('grid line param must be a number');
        }
        if ($gridLine < 0 or $gridLine > 3) {
            throw new ParamErrorException('grid line param error!');
        }
        $this->excel->gridline($gridLine);
        return $this;
    }

    /**
     * setPrintedPortrait
     * @return $this
     * @throws Exception
     */
    private function setPortrait()
    {
        $this->excel->setPrintedPortrait();
        return $this;
    }

    /**
     * setPrintedLandscape
     * @return $this
     * @throws Exception
     */
    private function setLandscape()
    {
        $this->excel->setPrintedLandscape();
        return $this;
    }

    /**
     * Get the memory-mode excel instance
     * @param $fileName
     * @return $this
     */
    private function memory($fileName)
    {
        $this->excel->constMemory($fileName);
        return $this;
    }

    /**
     * the necessary operation for excel building.
     * @param string $fileName
     * @return $this
     * @throws Exception
     */
    private function getExcelNormalInstance($fileName = '')
    {
        $this->excel()->fileName($fileName);
        return $this;
    }

    /**
     * get the \Vtiful\Kernel\Excel instance with param path
     * it's necessary.
     * @param string $path
     * @return $this
     * @throws Exception
     */
    private function excel()
    {
        if (!empty($this->path) && !is_writable($this->path)) {
            throw new Exception('the path specified cannot be written.');
        }

        $this->excel = new \Vtiful\Kernel\Excel(['path' => $this->getPath()]);
        return $this;
    }

    public function getData()
    {
        return $this->data;
    }

    /**
     * get header from excel
     * @return mixed
     */
    public function getHeader()
    {
        return $this->header;
    }

    /**
     * set header for excel
     * @param $header
     * @return $this
     * @throws Exception
     */
    private function setHeader($header)
    {
        $this->header = $header;
        $this->excel->header($header);
        return $this;
    }

    /**
     * set protection for excel
     * @param string $password
     * @return $this
     */
    public function setProtection($password = '')
    {
        $this->protection = $password;
        return $this;
    }

    protected function setPrintMode($mode)
    {
        if ($mode === \Lysice\XlsWriter\Excel::MODE_PRINT_PORTRAIT) {
            $this->printMode = \Lysice\XlsWriter\Excel::MODE_PRINT_PORTRAIT;
        } else if ($mode === \Lysice\XlsWriter\Excel::MODE_PRINT_LANDSCAPE) {
            $this->printMode = \Lysice\XlsWriter\Excel::MODE_PRINT_LANDSCAPE;
        } else {
            throw new ConfigNotFoundException('Print mode ' . $mode . ' doesn\'t support!');
        }
    }

    public function modeRow($data)
    {
        foreach ($data as $index => $column) {
            $row = new Row($this->excel, $column, $index + 1, $this->rowFormats);
            $row->insert();
        }

        return $this;
    }

    /**
     * set data for excel export
     * @param $data
     * @return $this
     */
    private function modeData($data)
    {
        $this->data = $data;
        $this->excel->data($data);
        return $this;
    }

    /**
     * export the excel to path file
     * @return string
     */
    private function output()
    {
        $this->filePath = $this->excel->output();
        return $this->filePath;
    }

    /**
     * apply format to writer
     * @param WithDefaultFormat $export
     * @return $this
     * @throws Exceptions\FormatParamErrorException
     */
    public function setDefaultFormat(WithDefaultFormat $export)
    {
        $formats = $export->defaultFormats();
        if (count($formats) == 1) {
            $this->defaultContentFormat = $formats[0];
        } else if (count($formats) == 2) {
            $this->defaultHeaderFormat = $formats[0];
            $this->defaultContentFormat = $formats[1];
        } else if (count($formats) == 0){
            return $this;
        } else {
            $this->defaultHeaderFormat = $formats[0];
            $this->defaultContentFormat = $formats[1];
        }

        return $this;
    }

    /**
     * set the default format for content or header
     * tips:this operation will be used globally
     * @param $format DefaultFormat
     * @return $this
     * @throws Exceptions\FormatParamErrorException
     */
    public function defaultFormat(DefaultFormat $format)
    {
        $this->excel->defaultFormat($format->getFormatResource($this->excel->getHandle()));
        return $this;
    }
}
