<?php
//call the autoload
require 'vendor/autoload.php';
//load phpspreadsheet class using namespaces
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;

//cargar el template
//cargar desde el template xlsx
$reader = IOFactory::createReader('xlsx');
$spreadsheet = $reader->load("template.xlsx");
//añadir el contenido
//coneccion a base de datos
$connection = mysqli_connect('localhost','root', '', 'phpexceltest');
if(!connection) {
	exit('database error');
}
$data = mysqli_query($connection, "SELECT * FROM students");
// recorrer los datos
$contentStartRow = 3;
$currentContentRow = 3;
while ($item = mysql_fetch_array($data)) {
	//insertamos una fila despues de la fila actual(antes de la actual fila + 1)
	$spreadsheet->getActiveSheet()->insertNewRowBefore($currentContentRow + 1, 1);

	// llenamos la celda con los datos (data)
	$spreadsheet->getActiveSheet()
				->setCellValue('A'.$currentContentRow, $item['id'])
				->setCellValue('B'.$currentContentRow, $item['first_name'])
				->setCellValue('C'.$currentContentRow, $item['last_name'])
				->setCellValue('D'.$currentContentRow, $item['email'])
				->setCellValue('E'.$currentContentRow, $item['gender'])
				->setCellValue('F'.$currentContentRow, $item['class'])
				->setCellValue('G'.$currentContentRow, $item['score']);

	// incrementa el numero de fila actual
	$currentContentRow++;
}

//remover filas
$spreadsheet->getActiveSheet()->removeRow($currentContentRow, 2);

//condiciones de formateo
$condition = new Conditional();
//condicion
$condition->setConditionType(Conditional::CONDITION_CELLIS)
			->setOperatorType(Conditional::OPERATOR_LESSTHAN)
			->addCondition(70);

// estilos de condicion
$condition->getStyle()->getFill()->setFillType(Fill::FILL_SOLID)
			->getEndColor()->setARGB(Color::COLOR_RED);
$condition->getStyle()->getFont()->getColor()->setARGB(Color::COLOR_YELLOW);

//aplicamos la condicional dentro del rango de celdas
$contentEndRow = $currentContentRow - 1;
$conditionalStyles = $spreadsheet->getActiveSheet()
					->getStyle('G' . $contentStartRow . ':G' . $contentEndRow)
					->getConditionalStyles();

array_push($conditionalStyles, $condition);
$spreadsheet->getActiveSheet()
			->getStyle('G' . $contentStartRow . ':G' . $contentEndRow)
			->setConditionalStyles($conditionalStyles);

//crea el archivo excel con su nombre y lo guarda en el servidor
// $writer = new Xlsx($spreadsheet);
// $writer->save('students_table.xlsx');
// echo "archivo students_table creado";