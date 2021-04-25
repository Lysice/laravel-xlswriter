<?php

namespace Lysice\XlsWriter\Supports\Chart;

interface ChartInterface {
    public function draw(array $data) : Chart;
}
