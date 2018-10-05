<?php

use SunDrop\ExcelExplorer;

$loader = require_once __DIR__ . '/vendor/autoload.php';
$loader->addPsr4('SunDrop\\', __DIR__ . '/src');

$explorer = new ExcelExplorer();
$fp = \fopen('result.csv', 'w');
\fputcsv($fp, ['Row Numbers', 'Execution Seconds', '.csv', '.csv.zip', '.xls', '.xls.zip', '.xlsx', '.xlsx.zip']);
foreach ($explorer->getFilesSize() as $data) {
    \fputcsv($fp, \array_values($data));
}
\fclose($fp);