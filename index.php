<?php
/* LIMPIANDO DIRECTORIO EXPORTS */
$DIRECTORIO = "exports/";
$files = glob($DIRECTORIO . '*');
foreach ($files as $file) {
    if (is_file($file) && !in_array(basename($file), ['.htaccess', '.gitkeep'])) {
        unlink($file);
    }
}

/** DEFINIENDO EL DIRECTORIO DE LA PLANTILLA EXCEL */
$templatePath = 'template/template.xlsx';

/** NOMBRE DEL ARCHIVO */
$filename = 'exports/NombreExcel.xlsx';

/** COMPROBAR LA VERSION DE PHP */
$isPhp8OrHigher = version_compare(PHP_VERSION, '8.0.0', '>=');

if ($isPhp8OrHigher) {
    /** PHP 8: UTILIZAR PHPSPREADSHEET */
    require_once 'class/PHPSpreadSheet/autoload.php';
    $excelClass = 'PhpOffice\PhpSpreadsheet\Spreadsheet';

    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load($templatePath);
    $sheet = $spreadsheet->getActiveSheet();

    /** ESCRIBIR EN EL EXCEL */
    $sheet->setCellValue('A2', 'Prueba con PHP8 o superior');
    $sheet->setSelectedCell('A1');

    /* GUARDANDO EL ARCHIVO CREADO */
    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    $writer->save($filename);
} else {
    /** PHP 7 O ANTERIOR - USAR PHPEXCEL */
    require_once 'class/PHPExcel/PHPExcel.php';

    $objPHPExcel = PHPExcel_IOFactory::load($templatePath);
    $objPHPExcel->setActiveSheetIndex(0);

    /** ESCRIBIR EN EL EXCEL */
    $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Prueba con PHP7 o anterior');
    $objPHPExcel->getActiveSheet()->setSelectedCell('A1');

    /* GUARDANDO EL ARCHIVO CREADO */
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save($filename);
}

/** VERIFICAR SI EL ARCHIVO EXISTE Y DESCARGARLO */
if (file_exists($filename) && is_readable($filename)) {
    $size = filesize($filename);
    $name = basename($filename);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment; filename=\"$name\"");
    header('Content-Length: ' . $size);
    header('Cache-Control: max-age=0');

    $file = fopen($filename, 'rb');
    if ($file) {
        while (!feof($file) && connection_status() == 0) {
            echo fread($file, 1024 * 8);
            flush();
        }
        fclose($file);
    }
    exit;
} else {
    header('HTTP/1.0 404 Not Found');
    echo "El archivo no está disponible para descarga.";
}

?>