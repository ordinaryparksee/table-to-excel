<?php

require __DIR__.'/../vendor/autoload.php';

ini_set('display_errors', true);

use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use Vaxy\TableToExcel\TableParser;

$spreadsheet = TableParser::parseFromFile(__DIR__.'/sample.html', [
    'variable' => 100
]);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="sample.xlsx"');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
