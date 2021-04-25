<?php

namespace Lysice\XlsWriter\Interfaces;

interface ExportInterface
{
    public function download($export, string $fileName, string $writerType = '', array $headers = []);

    public function store($export, string $filePath, string $disk = null, string $writerType = null, $diskOptions = []);

    public function queue($export, string $filePath, string $disk = null, string $writerType = null, $diskOptions = []);

    public function raw($export, string $writerType);
}
