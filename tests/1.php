<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\Cache\Adapter as CacheAdapter;
use Symfony\Component\Cache\Psr16Cache as Cache;

error_reporting(-1);
ini_set('display_errors', 'On');

require_once __DIR__ . '/../vendor/autoload.php';

$cacheAdapter = new CacheAdapter\ArrayAdapter();
$cache = new Cache($cacheAdapter);

$euroFormatter = new NumberFormatter('nl_NL', NumberFormatter::CURRENCY);
$template = new Langmans\SpreadsheetTemplate\Template($cache);
$template->setFormatter('euro', fn($value) => is_numeric($value) ? $euroFormatter->format($value) : $value);
$spreadsheet = $template->load(__DIR__ . '/1.ods', [
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
                'salary' => null
            ],
            [
                'name' => 'Jane Doe',
                'age' => 36,
                'job' => 'Manager',
                'salary' => 1500
            ],
        ],
    ],
]);

$writer = IOFactory::createWriter($spreadsheet, IOFactory::WRITER_ODS);
header('Content-Disposition: attachment;filename="export.ods"');
$writer->save('php://output');
die;
