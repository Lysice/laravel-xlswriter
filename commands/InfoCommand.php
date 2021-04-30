<?php

namespace Lysice\XlsWriter\Commands;

use Illuminate\Console\Command;

class InfoCommand extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'xls:status';

    /**
     * @var string
     */
    protected $version = '1.0';

    /**
     * @var string
     */
    protected $docsUrl = 'https://github.com/Lysice/laravel-xlswriter';
    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'show status for laravel-xlswriter';

    /**
     * Create a new command instance.
     *
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        $result = [
            'loaded' =>  extension_loaded('xlswriter') ? 'yes' : 'no',
            'xlsWriter author' => function_exists('xlswriter_get_author') ? xlswriter_get_author() : '',
        ];
        $execInfo = shell_exec('php --ri xlswriter');
        $execInfo =  explode(PHP_EOL, $execInfo);
        foreach ($execInfo as $index => $item) {
            if (empty($item) or strpos($item, '=>') == false) {
                unset($execInfo[$index]);
            }
        }
        foreach ($execInfo as $value) {
            $arr = explode('=>', $value);
            $result[trim($arr[0])] = trim($arr[1]);
        }

        $data = [
            'version' => $this->version,
            'author' => 'lysice<https://github.com/Lysice>',
            'docs' => $this->docsUrl
        ];
        $this->displayTables('laravel-xlsWriter info:', $data);
        $this->displayTables('XlsWriter extension status:', $result);
    }

    /**
     * Display info tables.
     *
     * @param $data
     */
    protected function displayTables($title, $data)
    {
        $this->line($title);
        $this->table([], $this->parseTable($data));
    }

    /**
     * Make up the table for console display.
     *
     * @param $input
     *
     * @return array
     */
    protected function parseTable($input)
    {
        $input = (array) $input;

        return array_map(function ($key, $value) {
            return [
                'key'       => $key,
                'value'     => $value,
            ];
        }, array_keys($input), $input);
    }
}
