<?php
// Activamos el autoloader de Composer
require __DIR__.'/vendor/autoload.php';
// require_once 'bootstrap.php';

use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\Font;

// Instancia phpWord
// $phpWord = new PhpWord();
$documento = new PhpWord();

// Nueva sección
$seccion = $documento->addSection();

// Texto sin formato
$seccion->addText(
    htmlspecialchars(
        'Primer texto - Texto sin formato'
    )
);


// Guardando documento
// $objWriter = IOFactory::createWriter($documento, "Word2007");
$documento->save("Documento01.docx", "Word2007");

header("Content-Disposition: attachment; filename=Documento01.docx");
echo file_get_contents('Documento01.docx');
