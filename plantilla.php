<?php
// Activamos el autoloader de Composer
// require __DIR__.'/vendor/autoload.php';
require_once 'bootstrap.php';

use PhpOffice\PhpWord\TemplateProcessor;

$templateWord = new TemplateProcessor('plantilla.docx');
 
$nombre = "Sandra S.L.";
$direccion = "Mi dirección";
$municipio = "Mrd";
$provincia = "Bdj";
$cp = "02541";
$telefono = "24536784";
$texto = "TEXTO DE PRUEBA";
$texto = ucwords(strtolower($texto));


// --- Asignamos valores a la plantilla
$templateWord->setValue('nombre_empresa',$nombre);
$templateWord->setValue('direccion_empresa',$direccion);
$templateWord->setValue('municipio_empresa',$municipio);
$templateWord->setValue('provincia_empresa',$provincia);
$templateWord->setValue('cp_empresa',$cp);
$templateWord->setValue('telefono_empresa',$telefono);
$templateWord->setValue('mi_texto',$texto);

// --- Guardamos el documento
$templateWord->saveAs('Documento02.docx');

header("Content-Disposition: attachment; filename=Documento02.docx; charset=iso-8859-1");
echo file_get_contents('Documento02.docx');