<?php

use Lysice\XlsWriter\Excel;

return [
    /**
     *  type of document
     * currently we support two type:
     * Excel::XLSX
     * Excel::CSV
     */
    'extension' => [
        'xlsx'     => Excel::XLSX,
        'csv'      => Excel::CSV
    ],
    /**
     * the mode for xlswriter
     * support:
     * Excel::MODE_NORMAL
     * Excel::MODE_MEMORY
     */
    'mode' => Excel::MODE_NORMAL,

    /**
     * laravel-xlswriter's data export mode
     * we support two modes:
     * Excel::EXPORT_MODE_DATA, we will export all data
     */
    'export_mode' => Excel::EXPORT_MODE_DATA,
];
