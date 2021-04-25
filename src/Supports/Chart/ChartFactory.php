<?php

namespace Lysice\XlsWriter\Supports\Chart;

use Lysice\XlsWriter\Supports\Constants;

/**
 * Class ChartFactory
 * @package Lysice\XlsWriter\Supports\Chart
 */
class ChartFactory {
    /**
     * get charts array
     * @param $fileHandle
     * @param array $chartConfigs
     * @return array
     * @throws \Lysice\XlsWriter\Exceptions\ChartParamErrorException
     */
    public static function makeCharts($fileHandle, $chartConfigs = [])
    {
        $charts = [];
        foreach ($chartConfigs as $chartConfig) {
            $charts[] = static::make($fileHandle, $chartConfig['type'], $chartConfig);
        }

        return $charts;
    }

    /**
     * make the chart instance
     * @param $fileHandle
     * @param $type
     * @param $data
     * @return Chart
     * @throws \Lysice\XlsWriter\Exceptions\ChartParamErrorException
     */
    public static function make($fileHandle, $type, $data)
    {
        switch($type) {
            case Constants::CHART_AREA:
            case Constants::CHART_AREA_STACKED:
            case Constants::CHART_AREA_PERCENT:
                $chart = (new AreaChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_BAR:
            case Constants::CHART_BAR_STACKED:
            case Constants::CHART_BAR_STACKED_PERCENT:
                $chart = (new BarChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_COLUMN:
            case Constants::CHART_COLUMN_STACKED:
            case Constants::CHART_COLUMN_STACKED_PERCENT:
                $chart = (new ColumnChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_DOUGHNUT:
                $chart = (new DoughnutChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_LINE:
                $chart = (new LineChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_PIE:
                $chart = (new PieChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_SCATTER:
            case Constants::CHART_SCATTER_STRAIGHT:
            case Constants::CHART_SCATTER_STRAIGHT_WITH_MARKERS:
            case Constants::CHART_SCATTER_SMOOTH:
            case Constants::CHART_SCATTER_SMOOTH_WITH_MARKERS:
                $chart = (new ScatterChart($fileHandle, $type))->draw($data);
                break;
            case Constants::CHART_RADAR:
            case Constants::CHART_RADAR_WITH_MARKERS:
            case Constants::CHART_RADAR_FILLED:
                $chart = (new RadarChart($fileHandle, $type))->draw($data);
                break;
            default:
                $chart = (new NoneChart($fileHandle, $type))->draw($data);
                break;
        }
        return $chart;
    }
}
