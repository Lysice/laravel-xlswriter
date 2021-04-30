<?php

namespace Lysice\XlsWriter;

use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Support\Collection;
use Lysice\XlsWriter\Exceptions\TypeNotSupportException;
use Lysice\XlsWriter\FileSystem\FileSystem;
use Lysice\XlsWriter\FileSystem\TemporaryFile;
use Lysice\XlsWriter\Interfaces\ExportInterface;
use Lysice\XlsWriter\Interfaces\ImportInterface;

/**
 * Class Excel
 * @package Lysice\XlsWriter
 */
class Excel implements ExportInterface, ImportInterface
{
    const MODE_MEMORY = 1;
    const MODE_NORMAL = 0;

    const EXPORT_MODE_DATA = 1;
    const EXPORT_MODE_ROW = 2;

    /**
     * print mode landscape
     * @var integer
     */
    const MODE_PRINT_LANDSCAPE = 1;

    /**
     * print mode portrait
     * @var integer
     */
    const MODE_PRINT_PORTRAIT = 2;

    /**
     * xlsx type name for export
     * @var string
     */
    const TYPE_XLSX = 'xlsx';


    /**
     * csv type name for export
     * @var string
     */
    const TYPE_CSV = 'csv';

    /**
     * export xlsx type
     * @var string
     */
    const XLSX = 'Xlsx';

    /**
     * export csv type
     * @var string
     */
    const CSV  = 'Csv';

    /**
     * @var XlsWriter
     */
    protected $xlsWriter;

    /**
     * @var
     */
    protected $filesystem;

    /**
     * @var
     */
    protected $exportMode;

    protected $config;
    public function __construct(
        XlsWriter $xlsWriter,
        FileSystem $filesystem
    )
    {
        $this->xlsWriter = $xlsWriter;
        $this->filesystem = $filesystem;
        $this->config = config('xlswriter');
    }

    /**
     * @param $export
     * @param string $fileName
     * @param string $writerType
     * @param array $headers
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse
     * @throws Exceptions\Exception
     * @throws TypeNotSupportException
     */
    public function download($export, string $fileName = 'demo', string $writerType = 'xlsx', array $headers = [])
    {
        $exportedFile = $this->export($export, $fileName, $writerType);

        return response()->download(
            $exportedFile->getLocalPath(),
            $fileName . '.' . $writerType,
            $headers
        )->deleteFileAfterSend(true);
    }

    /**
     * @param $export
     * @param string $filePath
     * @param string|null $diskName
     * @param string|null $writerType
     * @param array $diskOptions
     * @return bool
     * @throws Exceptions\Exception
     * @throws TypeNotSupportException
     */
    public function store($export, string $filePath, string $diskName = null, string $writerType = null, $diskOptions = [])
    {
        // 队列化
//        if ($export instanceof ShouldQueue) {
//            return $this->queue($export, $filePath, $diskName, $writerType, $diskOptions);
//        }
        $extensions = explode('.', $filePath);
        if (count($extensions) != 2){
            throw new \InvalidArgumentException('the name for document is not a valid name! for example:demo.xlsx');
        }
        if (!key_exists($extensions[1], $this->config['extension'])) {
            throw new \InvalidArgumentException('extension doesnot support!');
        }
        if (empty($diskName)) {
            throw new \InvalidArgumentException('the name for document must be passed, null given');
        }
        if (empty($writerType)) {
            throw new \InvalidArgumentException('the export type for document must be passed, null given');
        }

        $temporaryFile = $this->export($export, $filePath, $writerType);

        $exported = $this->filesystem->disk($diskName, $diskOptions)->copy(
            $temporaryFile,
            $filePath
        );

        $temporaryFile->delete();

        return $exported;
    }

    public function queue($export, string $filePath, string $disk = null, string $writerType = null, $diskOptions = [])
    {
        // TODO: Implement queue() method.
    }

    public function raw($export, string $writerType)
    {
        // TODO: Implement raw() method.
    }

    public function import($import, $filePath, string $disk = null, string $readerType = null)
    {
        // TODO: Implement import() method.
    }

    public function toArray($import, $filePath, string $disk = null, string $readerType = null): array
    {
        // TODO: Implement toArray() method.
    }

    public function toCollection($import, $filePath, string $disk = null, string $readerType = null): Collection
    {
        // TODO: Implement toCollection() method.
    }

    public function queueImport(ShouldQueue $import, $filePath, string $disk = null, string $readerType = null)
    {
        // TODO: Implement queueImport() method.
    }

    /**
     * @param $export
     * @param string $fileName
     * @param string|null $writerType
     * @param array $headers
     * @return XlsWriter|mixed
     * @throws Exceptions\Exception
     * @throws TypeNotSupportException
     */
    protected function export($export, string $fileName = 'demo', string $writerType = 'xlsx') : TemporaryFile
    {
        $writerType = $this->detectType($writerType);
        return $this->xlsWriter->export($export, $writerType, $fileName);
    }

    /**
     * @param $type
     * @return mixed
     * @throws TypeNotSupportException
     */
    private function detectType($type)
    {
        if (!key_exists($type, $this->config['extension'])) {
            throw new TypeNotSupportException();
        }

        return $type;
    }
}
