<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\Cache\Adapter as CacheAdapter;
use Symfony\Component\Cache\Psr16Cache as Cache;

error_reporting(-1);
ini_set('display_errors', 'On');

require_once __DIR__ . '/../vendor/autoload.php';

$cacheAdapter = new CacheAdapter\ArrayAdapter();
$cache = new Cache($cacheAdapter);

$template = new Langmans\PhpSpreadsheet\Template($cache);
$spreadsheet = $template->load(__DIR__ . '/1.xlsx', [
    'company' => [
        'name' => 'Acme Inc',
        'address' => [
            'street' => 'Main St',
            'house_number' => '123',
            'house_number_addendum' => 'Apt 1',
            'city' => 'Annie-town',
            'state' => 'CA',
            'zip' => '12345',
        ],
        'employees' => [
            [
                'name' => 'John Doe',
                'age' => 42,
                'job' => 'Developer',
            ],
            [
                'name' => 'Jane Doe',
                'age' => 36,
                'job' => 'Manager',
            ],
        ],
    ],
]);
dd($spreadsheet->getActiveSheet()->toArray());

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="export.xlsx"');
$writer->save('php://output');
die;