<?php

namespace Lysice\XlsWriter\Traits;

use Illuminate\Foundation\Bus\PendingDispatch;
use Lysice\XlsWriter\Exceptions\Exception;
use Lysice\XlsWriter\Exceptions\ParamErrorException;
use Lysice\XlsWriter\Interfaces\ExportInterface;

trait Exportable
{
    /**
     * @param string      $fileName
     * @param string|null $writerType
     * @param array       $headers
     *
     * @throws ParamErrorException
     * @return \Illuminate\Http\Response|\Symfony\Component\HttpFoundation\BinaryFileResponse
     */
    public function download(string $fileName = null, string $writerType = 'xlsx', array $downloadHeaders = [])
    {
        $headers    = empty($downloadHeaders) ? ($this->downloadHeaders ?? []) : $downloadHeaders;
        $fileName   = empty($fileName) ? ($this->fileName ?? '') : $fileName;
        $writerType = empty($writerType) ? ($this->writerType ?? 'xlsx') : $writerType;

        if (null === $fileName) {
            throw new ParamErrorException('filename required for download');
        }
        if (null === $writerType) {
            throw new ParamErrorException('writerType required for download');
        }

        return $this->getExporter()->download($this, $fileName, $writerType, $headers);
    }

    /**
     * @param string|null $filePath
     * @param string $fileName
     * @param string|null $disk
     * @param string|null $writerType
     * @param array $diskOptions
     * @return mixed
     */
    public function store(string $filePath = null, string $disk = null, string $writerType = null, $diskOptions = [])
    {
        $filePath = $filePath ?? $this->filePath ?? null;

        if (null === $filePath) {
            throw ParamErrorException::export();
        }
        if (null === $disk) {
            throw new ParamErrorException('disk required for store');
        }
        if (null === $writerType) {
            throw new ParamErrorException('writerType required for store');
        }

        return $this->getExporter()->store(
            $this,
            $filePath,
            $disk ?? $this->disk ?? null,
            $writerType ?? $this->writerType ?? null,
            $diskOptions ?? $this->diskOptions ?? []
        );
    }

    /**
     * @param string|null $filePath
     * @param string|null $disk
     * @param string|null $writerType
     * @param mixed       $diskOptions
     *
     * @throws ParamErrorException
     * @return PendingDispatch
     */
    public function queue(string $filePath = null, string $disk = null, string $writerType = null, $diskOptions = [])
    {
        $filePath = $filePath ?? $this->filePath ?? null;

        if (null === $filePath) {
            throw ParamErrorException::export();
        }

        return $this->getExporter()->queue(
            $this,
            $filePath,
            $disk ?? $this->disk ?? null,
            $writerType ?? $this->writerType ?? null,
            $diskOptions ?? $this->diskOptions ?? []
        );
    }

    /**
     * @param string|null $writerType
     *
     * @return string
     */
    public function raw($writerType = 'xlsx')
    {
        return $this->getExporter()->raw($this, $writerType);
    }

    /**
     * Create an HTTP response that represents the object.
     * @param $request
     * @return \Illuminate\Http\Response|\Symfony\Component\HttpFoundation\BinaryFileResponse
     */
    public function toResponse($request)
    {
        return $this->download();
    }

    /**
     * @return ExportInterface
     */
    private function getExporter(): ExportInterface
    {
        return app(ExportInterface::class);
    }
}
