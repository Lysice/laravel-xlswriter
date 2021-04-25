<?php

namespace Lysice\XlsWriter\FileSystem;

class ExportedFile extends TemporaryFile{

    private $filePath;

    public function __construct(string $filePath)
    {
        $this->filePath = realpath($filePath);
    }

    /**
     * @inheritDoc
     */
    public function getLocalPath(): string
    {
        return $this->filePath;
    }

    /**
     * @inheritDoc
     */
    public function exists(): bool
    {
        return file_exists($this->filePath);
    }

    /**
     * @inheritDoc
     */
    public function put($contents)
    {
        file_put_contents($this->filePath, $contents);
    }

    /**
     * @inheritDoc
     */
    public function delete(): bool
    {
        if (@unlink($this->filePath) || !$this->exists()) {
            return true;
        }

        return unlink($this->filePath);
    }

    /**
     * @inheritDoc
     */
    public function readStream()
    {
        return fopen($this->getLocalPath(), 'rb+');
    }

    /**
     * @inheritDoc
     */
    public function contents(): string
    {
        return file_get_contents($this->filePath);
    }
}
