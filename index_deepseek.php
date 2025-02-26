<?php
// Definir constantes para rutas
define('DIRECTORIO_EXPORTS', __DIR__ . '/exports/');
define('TEMPLATE_PATH', __DIR__ . '/template/template.xlsx');
define('FILENAME', DIRECTORIO_EXPORTS . 'ExcelGenerado.xlsx');

// Limpiar directorio exports
if (!is_dir(DIRECTORIO_EXPORTS)) {
    die("El directorio " . DIRECTORIO_EXPORTS . " no existe.");
}

$files = glob(DIRECTORIO_EXPORTS . '*', GLOB_NOSORT);
array_map('unlink', array_filter($files, function ($file) {
    return is_file($file) && !in_array(basename($file), ['.htaccess', '.gitkeep']);
}));

// Comprobar la versión de PHP
$isPhp8OrHigher = version_compare(PHP_VERSION, '8.0.0', '>=');

if ($isPhp8OrHigher) {
    // PHP 8: Utilizar PHPSpreadsheet
    require_once __DIR__ . '/class/PHPSpreadSheet/autoload.php';
    $excelClass = 'PhpOffice\PhpSpreadsheet\Spreadsheet';

    try {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load(TEMPLATE_PATH);
        $sheet = $spreadsheet->getActiveSheet();

        // Escribir en el Excel
        $sheet->setCellValue('A2', 'Prueba con PHP8 o superior');
        $sheet->setSelectedCell('A1');

        // Guardar el archivo creado
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save(FILENAME);
    } catch (\Exception $e) {
        die("Error al procesar la plantilla: " . $e->getMessage());
    }
} else {
    // PHP 7 o anterior - Usar PHPExcel
    require_once __DIR__ . '/class/PHPExcel/PHPExcel.php';

    try {
        $objPHPExcel = PHPExcel_IOFactory::load(TEMPLATE_PATH);
        $objPHPExcel->setActiveSheetIndex(0);

        // Escribir en el Excel
        $objPHPExcel->getActiveSheet()->setCellValue('A2', 'Prueba con PHP7 o anterior');
        $objPHPExcel->getActiveSheet()->setSelectedCell('A1');

        // Guardar el archivo creado
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(FILENAME);
    } catch (\Exception $e) {
        die("Error al procesar la plantilla: " . $e->getMessage());
    }
}

// Verificar si el archivo existe y descargarlo
if (file_exists(FILENAME) && is_readable(FILENAME)) {
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment; filename=\"" . basename(FILENAME) . "\"");
    header('Content-Length: ' . filesize(FILENAME));
    readfile(FILENAME);
    exit;
} else {
    header('HTTP/1.0 404 Not Found');
    echo "El archivo no está disponible para descarga.";
}
