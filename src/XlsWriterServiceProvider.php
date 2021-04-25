<?php

namespace Lysice\XlsWriter;

use Lysice\XlsWriter\Commands\ExportMakeCommand;
use Lysice\XlsWriter\Commands\InfoCommand;
use Lysice\XlsWriter\FileSystem\FileSystem;
use Lysice\XlsWriter\Interfaces\ExportInterface;

class XlsWriterServiceProvider extends \Illuminate\Support\ServiceProvider
{
    protected $commands = [
        InfoCommand::class,
        ExportMakeCommand::class
    ];

    public function boot()
    {
        if ($this->app->runningInConsole()) {
            $this->publishes([
                __DIR__.'/../config/xlswriter.php' => config_path('xlswriter.php'),
            ], 'config');
        }
    }

    public function register()
    {
        $this->app->singleton(XlsWriter::class, function () {
            return new XlsWriter();
        });
        $this->app->singleton(FileSystem::class, function () {
            return new FileSystem($this->app->make('filesystem'));
        });
        $this->app->singleton(Excel::class, function (){
            return new Excel(
                app()->make(XlsWriter::class),
                $this->app->make(Filesystem::class)
            );
        });
        $this->app->alias(Excel::class, ExportInterface::class);
        $this->app->alias(Excel::class, ExportInterface::class);
        $this->registerCommands();
    }

    /**
     * register the commands
     */
    private function registerCommands()
    {
        $this->commands($this->commands);
    }
}
