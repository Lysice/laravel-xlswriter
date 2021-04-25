<?php

if(!function_exists('getTmpDir')) {
    function getTmpDir(): string
    {
        $tmp = ini_get('upload_tmp_dir');

        if ($tmp !== False && file_exists($tmp)) {
            return realpath($tmp);
        }

        return realpath(sys_get_temp_dir());
    }
}
if (!function_exists('version_check')) {
    function versionCheck($checkVersion)
    {
        if (function_exists('xlswriter_get_version')) {
            $version = versionToInt(xlswriter_get_version());
            $current = versionToInt($checkVersion);
            return $version == $current;
        } else {
            throw new LogicException('xlswriter_get_version function doesnot exists! please check your xlswriter extension first');
        }
    }
}
function versionToInt($version)
{
    $vArr = explode('.', $version);
    $total = 0;
    foreach ($vArr as $item) {
        $total = $total * 100 + $item;
    }

    return $total;
}
