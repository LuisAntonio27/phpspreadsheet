<?php
//call the autoload
require 'vendor/autoload.php';
//load phpspreadsheet class using namespaces
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call xlsx writer
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//phpspreadsheet Date class
use PhpOffice\PhpSpreadsheet\Shared\Date;
//phpspreadsheet numberformat style class
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
//rich text class
use \PhpOffice\PhpSpreadsheet\RichText\RichText;
//phpspreadsheet style color
use \PhpOffice\PhpSpreadsheet\Style\Color;

//make a new spreadsheet object
$spreadsheet = new Spreadsheet();
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

//set default font
$spreadsheet->getDefaultStyle()
			->getFont()
			->setName('Arial')
			->setSize(10);

//set column dimension to auto size
$sheet->getColumnDimension('B')
		->setAutoSize(true);
$sheet->getColumnDimension('C')
		->setAutoSize(true);

//simple text data
$sheet->setCellValue('A1',"String")
		->setCellValue('B1',"Simple Text")
		->setCellValue('C1',"This is Phpspreadsheet");

//symbols
$sheet->setCellValue('A2',"String")
		->setCellValue('B2',"Symbols")
		->setCellValue('C2',"ÚÔÛï¢£´°ƤǠњс҃ҭ");

//utf-8 string
$sheet->setCellValue('A3',"String")
		->setCellValue('B3',"UTF-8")
		->setCellValue('C3',"добро пожаловать в мой учебник видео");

//integer
$sheet->setCellValue('A4',"Number")
		->setCellValue('B4',"Integer")
		->setCellValue('C4',55);

//float
$sheet->setCellValue('A5',"Number")
		->setCellValue('B5',"Float")
		->setCellValue('C5',55.55);

//negative
$sheet->setCellValue('A6',"Number")
		->setCellValue('B6',"Negative")
		->setCellValue('C6',-55.55);
//boolean
$sheet->setCellValue('A7',"Number")
		->setCellValue('B7',"Boolean")
		->setCellValue('C7',true)
		->setCellValue('D7',false);

//date datatype
//make a variable from current timestamp
$dateTimeNow = time();
$hoy = date('d-m-Y');

//date
$sheet->setCellValue('A8',"Date/Time")
		->setCellValue('B8',"Date")
		->setCellValue('C8',Date::PHPToExcel($dateTimeNow))
		->setCellValue('D8',$hoy);

//set the cell format into a date
$sheet->getStyle('C8')
		->getNumberFormat()
		->setFormatCode(NumberFormat::FORMAT_DATE_YYYYMMDD2);


//date with time
$sheet->setCellValue('A9',"Date/Time")
		->setCellValue('B9',"Date Time")
		->setCellValue('C9',Date::PHPToExcel($dateTimeNow));

//set the cell format into a date
$sheet->getStyle('C9')
		->getNumberFormat()
		->setFormatCode(NumberFormat::FORMAT_DATE_DATETIME);

//only time
$sheet->setCellValue('A10',"Date/Time")
		->setCellValue('B10',"Only Time")
		->setCellValue('C10',Date::PHPToExcel($dateTimeNow));

//set the cell format into a date
$sheet->getStyle('C10')
		->getNumberFormat()
		->setFormatCode(NumberFormat::FORMAT_DATE_TIME4);

//rich text
$sheet->setCellValue('A11',"Rich text");

$richText = new RichText();
$richText->createText('normal text ');
$payable = $richText->createTextRun('bold italic and dark green');
$payable->getFont()->setBold(true);
$payable->getFont()->setItalic(true);
$payable->getFont()->setColor( new Color( Color::COLOR_DARKGREEN ) );

//add a rich text
$redText = $richText->createTextRun('red text');
$redText->getFont()->setColor( new Color( Color::COLOR_RED ) );

$richText->createText(' normal text again');
$sheet->getCell('C11')->setValue($richText);

//hyperlink
$sheet->setCellValue('A12',"Hyperlink")
		->setCellValue('B12',"Cell Hyperlink")
		->setCellValue('C12',"Visit Gemul's Channel");

//set the cell as hyperlink
$sheet->getCell('C12')
		->getHyperlink()
		->setUrl('https://www.youtube.com')
		->setTooltip('Ir a youtube');

//hyperlink with formula
$sheet->setCellValue('A13',"Hyperlink")
		->setCellValue('B13',"Formula Hyperlink")
		->setCellValue('C13',"=HYPERLINK(\"https://www.facebook.com\",\"Ir a facebook\")");

//change worksheet name
$sheet->setTitle('Phpspreadsheet');

//crea el archivo excel con su nombre y lo guarda en el servidor
$writer = new Xlsx($spreadsheet);
$writer->save('formatos_valores.xlsx');
echo "archivo formatos_valores creado";