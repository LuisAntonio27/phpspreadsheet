<?php

//llama el autoload
require 'vendor/autoload.php';

//carga la clase spreadsheet usando namespace
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//llama la clase xlsx writes para hacer un archivo de excel
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// crea un objeto spreadsheet
$spreadsheet = new Spreadsheet();
//obtiene la hoja actual
$sheet = $spreadsheet->getActiveSheet();
//configura el valor de la celda A1 en Hello World
$sheet->setCellValue('A1', 'Hello World !');

//crea el archivo excel con su nombre y lo guarda en el servidor
$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');
echo "hola mundo creado";