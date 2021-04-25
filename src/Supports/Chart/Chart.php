<?php

namespace Lysice\XlsWriter\Supports\Chart;

use Illuminate\Support\Facades\Log;
use Lysice\XlsWriter\Exceptions\ChartParamErrorException;
use Lysice\XlsWriter\Exceptions\ParamErrorException;
use Vtiful\Kernel\Excel;
use Lysice\XlsWriter\Supports\Constants;

class Chart implements ChartInterface {
    /**
     * @var \Vtiful\Kernel\Chart
     */
    protected $chart;

    /**
     * @var Excel
     */
    protected $fileHandle;

    /**
     * @var string
     */
    protected $chartType;

    /**
     * @var int
     */
    public $row;

    /**
     * @var int
     */
    public $column;

    protected function getChartTypes()
    {
        return [
//        \Vtiful\Kernel\Chart::CHART_NONE, // 无
            Constants::CHART_AREA, // 面积图
            Constants::CHART_AREA_STACKED,// 面积图-堆积
            Constants::CHART_AREA_PERCENT,// 面积图-堆积百分比
            Constants::CHART_BAR,// 条形图
            Constants::CHART_BAR_STACKED,// 条形图-堆积
            Constants::CHART_BAR_STACKED_PERCENT,// 条形图-堆积百分比
            Constants::CHART_COLUMN,// 直方图
            Constants::CHART_COLUMN_STACKED,// 直方图-堆积
            Constants::CHART_COLUMN_STACKED_PERCENT,// 直方图-堆积百分比
            Constants::CHART_DOUGHNUT,// 圆环图
            Constants::CHART_LINE,// 折线图
            Constants::CHART_PIE,// 饼图
//            Constants::CHART_SCATTER,// 散点图
//            Constants::CHART_SCATTER_STRAIGHT,// 散点图-直线
//            Constants::CHART_SCATTER_STRAIGHT_WITH_MARKERS,// 散点图-直线链接标记
//            Constants::CHART_SCATTER_SMOOTH,//散点图 - 平滑线
//            Constants::CHART_SCATTER_SMOOTH_WITH_MARKERS, // 散点图 - 平滑线链接标记
            Constants::CHART_RADAR, // 雷达图
            Constants::CHART_RADAR_WITH_MARKERS, //雷达图 - 带标记
            Constants::CHART_RADAR_FILLED
        ];
    }

    public function __construct($fileHandle, $chartType = '')
    {
        if (!is_resource($fileHandle)) {
            throw new ChartParamErrorException("FileHandle must be a resource of Excel");
        }

        if (empty($chartType)) {
            throw new ChartParamErrorException("chart type required");
        }

        if (!in_array($chartType, $this->getChartTypes())) {
            throw new ChartParamErrorException("Chart Type " . $chartType . 'doesn\'t support');
        }
        $this->fileHandle = $fileHandle;
        $this->chartType = $chartType;
    }

    public function create()
    {
        $this->chart = new \Vtiful\Kernel\Chart($this->fileHandle, $this->chartType);
        return $this;
    }

    /**
     * @param string $title
     * @return $this
     */
    protected function setTitle($title = 'title')
    {
        $this->chart->title($title);
        return $this;
    }

    /**
     * set the position where the chart be inserted.
     * @param int $row
     * @param int $column
     * @return $this
     */
    protected function setPosition($row = 1, $column = 1)
    {
        $this->row = $row;
        $this->column = $column;
        return $this;
    }

    /**
     * @param array $series
     * @return mixed
     * @throws ChartParamErrorException
     */
    public function setSeries(array $series = [])
    {
        if (empty($series)) {
            throw new ChartParamErrorException("Chart series data must be not empty");
        }

        foreach ($series as $item) {
            if (!$this->seriesCheck($item['data'])) {
                throw new ChartParamErrorException('series data error!');
            }

            if (isset($item['category'])) {
                $this->chart->series($item['data'], $item['category']);
            } else {
                $this->chart->series($item['data']);
            }

            if (isset($item['name'])) {
                $this->chart->seriesName($item['name']);
            }
        }

        return $this;
    }


    /**
     * chart style. you can refer to Excel 2007 "design" styles
     * @param $style
     * @return $this
     */
    protected function setStyle($style)
    {
        if ($style < 1 or $style > 48) {
            throw new ParamErrorException('AreaChart style must be in 1 ~ 48');
        }

        $this->chart->style($style);
        return $this;
    }

    /**
     * todo
     * @param $seriesData
     * @return bool
     */
    private function seriesCheck($seriesData) {
        // 1.工作表验证
        // 2.格式验证
        return true;
    }

    /**
     * return a resource of chart
     * @return mixed
     */
    public function toResource()
    {
        return $this->chart->toResource();
    }

    /**
     * named x axis
     * @param string $xName
     * @return $this
     */
    protected function setAxisNameX($xName = 'x axis')
    {
        $this->chart->axisNameX($xName);
        return $this;
    }

    /**
     * name y axis
     * @param string $yName
     * @return $this
     */
    protected function setAxisNameY($yName = 'y axis')
    {
        $this->chart->axisNameY($yName);
        return $this;
    }

    /**
     * @param array $data
     * @return Chart
     * @throws \Lysice\XlsWriter\Exceptions\ChartParamErrorException
     */
    public function draw(array $data): Chart
    {
        return $this->create()
            ->setTitle($data['title'])
            ->setStyle($data['style'])
//            ->setAxisNameX($data['xName'])
//            ->setAxisNameY($data['yName'])
            ->setPosition($data['row'], $data['column'])
            ->setSeries($data['series']);
    }
}
