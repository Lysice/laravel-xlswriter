<?php

namespace Lysice\XlsWriter\FileSystem;

use Illuminate\Contracts\Filesystem\Factory;
use Lysice\XlsWriter\FileSystem\Disk;

class FileSystem {
    private $filesystem;

    public function __construct(Factory $filesystem)
    {
        $this->filesystem = $filesystem;
    }

    /**
     * @param string|null $disk
     * @param array       $diskOptions
     *
     * @return Disk
     */
    public function disk(string $disk = null, array $diskOptions = []): Disk
    {
        return new Disk(
            $this->filesystem->disk($disk),
            $disk,
            $diskOptions
        );
    }
}
