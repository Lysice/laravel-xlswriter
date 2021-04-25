<?php

namespace Lysice\XlsWriter\Traits;

use Lysice\XlsWriter\Exceptions\DirectoryCannotWritableException;
use Lysice\XlsWriter\Exceptions\DirectoryNotExistsException;
use Lysice\XlsWriter\Exceptions\ExportFailedException;
use Lysice\XlsWriter\XlsWriter;

trait CommonTrait
{
    /**
     * set filename for the export file
     * it's necessary.
     * @param $fileName
     * @return $this
     */
    private function fileName($fileName)
    {
        $this->excel->fileName($fileName);
        return $this;
    }

    /**
     * Apply the callback's query changes if the given "value" is true.
     *
     * @param  mixed  $value
     * @param  callable  $callback
     * @param  callable|null  $default
     * @return mixed|$this
     */
    public function when($value, $callback, $default = null)
    {
        if ($value) {
            return $callback($this, $value) ?: $this;
        } elseif ($default) {
            return $default($this, $value) ?: $this;
        }

        return $this;
    }

    /**
     * @param string $name
     * @return mixed|string
     */
    private function config($name = '')
    {
        if (isset($this->config[$name])) {
            return $this->config[$name];
        }
        return '';
    }

    /**
     * get default dir
     * @return string
     */
    function getTmpDir(): string
    {
        $tmp = ini_get('upload_tmp_dir');

        if ($tmp !== False && file_exists($tmp)) {
            return realpath($tmp);
        }

        return realpath(sys_get_temp_dir());
    }

    function getPath() : string
    {
        if (!empty($this->path)) {
            if (!file_exists($this->path)) {
                throw new DirectoryNotExistsException('the path is not exist');
            }
            if (!is_writable($this->path)) {
                throw new DirectoryCannotWritableException('the path cannot be written');
            }
            if (file_exists($this->path) && is_writable($this->path)) {
                return realpath($this->path);
            }
        }

        return $this->getTmpDir();
    }
}
