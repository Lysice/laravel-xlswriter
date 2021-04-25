<?php
namespace Lysice\XlsWriter\Interfaces;

use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Support\Collection;

interface ImportInterface
{
    public function import($import, $filePath, string $disk = null, string $readerType = null);
    public function toArray($import, $filePath, string $disk = null, string $readerType = null): array;
    public function toCollection($import, $filePath, string $disk = null, string $readerType = null): Collection;
    public function queueImport(ShouldQueue $import, $filePath, string $disk = null, string $readerType = null);
}
