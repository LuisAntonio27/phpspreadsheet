<?php
//call the autoload
require 'vendor/autoload.php';
//load phpspreadsheet class using namespaces
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call iofactory instead of xlsx writer
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

//styling arrays
//table head style
$tableHead = [
	'font'=>[
		'color'=>[
			'rgb'=>'FFFFFF'
		],
		'bold'=>true,
		'size'=>11
	],
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '538ED5'
		]
	],
];
//even row
$evenRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00BDFF'
		]
	]
];
//odd row
$oddRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00EAFF'
		]
	]
];

//styling arrays end

//make a new spreadsheet object
$spreadsheet = new Spreadsheet();
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

//set default font
$spreadsheet->getDefaultStyle()
	->getFont()
	->setName('Arial')
	->setSize(10);

//heading
$sheet->setCellValue('A1',"Participant Students");

//merge heading
$sheet->mergeCells("A1:F1");

// set font style
$sheet->getStyle('A1')->getFont()->setSize(20);

// set cell alignment
$sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

//setting column width
$sheet->getColumnDimension('A')->setWidth(5);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(20);
$sheet->getColumnDimension('D')->setWidth(30);
$sheet->getColumnDimension('E')->setWidth(12);
$sheet->getColumnDimension('F')->setWidth(10);

//header text
$sheet->setCellValue('A2',"ID")
	->setCellValue('B2',"First Name")
	->setCellValue('C2',"Last Name")
	->setCellValue('D2',"Email")
	->setCellValue('E2',"Gender")
	->setCellValue('F2',"Class");

//set font style and background color
$sheet->getStyle('A2:F2')->applyFromArray($tableHead);

//the content
//read the json file
$file = file_get_contents('student-data.json');
$studentData = json_decode($file,true);

//loop through the data
//current row
$row=3;
foreach($studentData as $student){
	$sheet->setCellValue('A'.$row , $student['id'])
		->setCellValue('B'.$row , $student['first_name'])
		->setCellValue('C'.$row , $student['last_name'])
		->setCellValue('D'.$row , $student['email'])
		->setCellValue('E'.$row , $student['gender'])
		->setCellValue('F'.$row , $student['class']);

	//set row style
	if( $row % 2 == 0 ){
		//even row
		$sheet->getStyle('A'.$row.':F'.$row)->applyFromArray($evenRow);
	}else{
		//odd row
		$sheet->getStyle('A'.$row.':F'.$row)->applyFromArray($oddRow);
	}
	//increment row
	$row++;
}

//autofilter
//define first row and last row
$firstRow=2;
$lastRow=$row-1;
//set the autofilter
$sheet->setAutoFilter("A".$firstRow.":F".$lastRow);

//crea el archivo excel con su nombre y lo guarda en el servidor
$writer = new Xlsx($spreadsheet);
$writer->save('students_table.xlsx');
echo "archivo students_table creado";