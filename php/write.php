<?php
require("vendor/autoload.php");

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;

// 新規スプレットシートとしてnew
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// サンプルのエクセル通りにデータを作る
$sheet->setCellValue('A1', '2015/4/5 13:34');
$sheet->setCellValue('A2', '2015/4/5 3:41');
$sheet->setCellValue('A3', '2015/4/6 12:46');
$sheet->setCellValue('A4', '2015/4/8 8:59');
$sheet->setCellValue('A5', '2015/4/10 2:07');
$sheet->setCellValue('A6', '2015/4/10 18:10');
$sheet->setCellValue('A7', '2015/4/10 2:40');
$sheet->setCellValue('B1', 'Apples');
$sheet->setCellValue('B2', 'Cherries');
$sheet->setCellValue('B3', 'Pears');
$sheet->setCellValue('B4', 'Oranges');
$sheet->setCellValue('B5', 'Apples');
$sheet->setCellValue('B6', 'Bananas');
$sheet->setCellValue('B7', 'Strawberries');
$sheet->setCellValue('C1', 73);
$sheet->setCellValue('C2', 85);
$sheet->setCellValue('C3', 14);
$sheet->setCellValue('C4', 52);
$sheet->setCellValue('C5', 152);
$sheet->setCellValue('C6', 23);
$sheet->setCellValue('C7', 98);

// 最後の行にC列のSUM情報を追加する
$sheet->setCellValue('C8', '=SUM(C1:C7)');

// 保存
$writer = new Writer($spreadsheet);
$writer->save('./excel/php_write.xlsx');
