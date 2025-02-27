<?php
/* LIMPIANDO DIRECTORIO EXPORTS */
$DIRECTORIO = "exports/";
if (!is_dir($DIRECTORIO)) {
    mkdir($DIRECTORIO, 0777, true);
}
$files = glob($DIRECTORIO . '*');
foreach ($files as $file) {
    if (is_file($file)) {
        unlink($file);
    }
}

/** DEFINIENDO EL DIRECTORIO DE LA PLANTILLA EXCEL */
$templatePath = 'template/template.xlsx';
/** NOMBRE DEL ARCHIVO */
$filename = 'exports/NombreExcel.xlsx';

/** FUNCION PARA GENERAR EXCEL */
function generarExcel($templatePath, $filename, $versionPhp8) {
    if ($versionPhp8) {
        require_once 'class/PHPSpreadSheet/autoload.php';
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($templatePath);
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A2', 'Prueba con PHP8 o superior');
        $sheet->setSelectedCell('A1');
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($filename);
    } else {
        require_once 'class/PHPExcel/PHPExcel.php';
        $objPHPExcel = PHPExcel_IOFactory::load($templatePath);
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Prueba con PHP7 o anterior');
        $objPHPExcel->getActiveSheet()->setSelectedCell('A1');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save($filename);
    }
}

/** COMPROBAR LA VERSION DE PHP */
$isPhp8OrHigher = version_compare(PHP_VERSION, '8.0.0', '>=');

try {
    generarExcel($templatePath, $filename, $isPhp8OrHigher);
} catch (Exception $e) {
    header('HTTP/1.0 500 Internal Server Error');
    echo "Error al generar el archivo: " . $e->getMessage();
    exit;
}

/** VERIFICAR SI EL ARCHIVO EXISTE Y DESCARGARLO */
if (file_exists($filename) && is_readable($filename)) {
    $size = filesize($filename);
    $name = basename($filename);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment; filename=\"$name\"");
    header('Content-Length: ' . $size);
    header('Cache-Control: max-age=0');

    ob_clean();
    flush();
    readfile($filename);
    exit;
} else {
    header('HTTP/1.0 404 Not Found');
    echo "El archivo no está disponible para descarga.";
    exit;
}
?>