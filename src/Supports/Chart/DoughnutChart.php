<?php

namespace Lysice\XlsWriter\Supports\Chart;

class DoughnutChart extends Chart implements ChartInterface
{
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
            ->setPosition($data['row'], $data['column'])
            ->setSeries($data['series']);
    }
}
