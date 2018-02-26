<?php
require("vendor/autoload.php");

ini_set('date.timezone', 'Asia/Tokyo');
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;

// ファイル読み込み
$reader      = new Reader();
$spreadsheet = $reader->load('./excel/example.xlsx');

// アクティブなシートから全データを取得
$data = $spreadsheet->getActiveSheet()->toArray();

// 表示
foreach ($data as $val) {
    echo $val[0] . "\t" . $val[1] . "\t" . $val[2] . "\n";
}
