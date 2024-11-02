<?php
require 'vendor/autoload.php'; // Para cargar PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$phone = $_POST['phone'];
$date = $_POST['date'];
$startTime = $_POST['startTime'];
$endTime = $_POST['endTime'];

// Cargar el archivo Excel o crear uno nuevo si no existe
$filePath = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRQJMYm0PHKNunwKvlfqdbskbfUJOqaCBGaIXfzwaCVyIP-1lWqv5XNzRiHwN6VHyLmoZSCTuDEu8Q3/pub?output=xlsx';
if (file_exists($filePath)) {
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'Teléfono')
          ->setCellValue('B1', 'Fecha')
          ->setCellValue('C1', 'Hora de inicio')
          ->setCellValue('D1', 'Hora de fin');
}

// Seleccionar la hoja activa y agregar los datos
$sheet = $spreadsheet->getActiveSheet();
$row = $sheet->getHighestRow() + 1;
$sheet->setCellValue("A{$row}", $phone)
      ->setCellValue("B{$row}", $date)
      ->setCellValue("C{$row}", $startTime)
      ->setCellValue("D{$row}", $endTime);

// Guardar el archivo
$writer = new Xlsx($spreadsheet);
$writer->save($filePath);

echo "Datos guardados exitosamente";
?>
