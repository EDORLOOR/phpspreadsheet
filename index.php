<?php
require_once "phpspreadsheet/vendor/autoload.php";

$detalles = array(
    array(
        "0" => "CONTROL",
        "1" => "PESO",
        "2" => "PROPIEDAD",
        "3" => "PROPIETARIO"
    ),

    array(
        "0" => "602",
        "1" => "0,28770",
        "2" => "2-1602",
        "3" => "MA HANZHONG"
    ),

    array(
        "0" => "601",
        "1" => "1,21400",
        "2" => "1-704",
        "3" => "BANCOLOMBIA"
    ),

    array(
        "0" => "0",
        "1" => "98,43370000",
        "2" => "INASISTENTES",
        "3" => "INASISTENTES"
    ),
);

$filename = "Sunvote.xlsx";
header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheet‌​ml.sheet");
header('Content-Disposition: attachment; filename="' . $filename. '"');

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('plantilla.xlsx');

$worksheet = $spreadsheet->getActiveSheet();

$worksheet->setTitle("SunVote");

$i = 1;
$letras = 64;

foreach ($detalles as $x => $detalle) {
    if ($x) {
        $i++;
        $worksheet->insertNewRowBefore($i, 1);
    }
	$letras = 64;
	foreach ($detalle as $x2 => $valor) {
		$letras += 1;
		$worksheet->getCell(chr($letras).$i)->setValue($valor);
	}
}

$column = $worksheet->getColumnDimension("A");
$column->setAutoSize(true);

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save("php://output");
?>
