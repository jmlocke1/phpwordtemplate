<?php
// Activamos el autoloader de Composer
// require __DIR__.'/vendor/autoload.php';
require_once 'bootstrap.php';

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

// Texto con formato
$seccion->addText(
    htmlspecialchars(
            'Segundo texto con formato'
    ),
    array('name' => 'Arial', 'size' => '12', 'bold' => 'true')
);

// Texto con fuente personalizada
$fuente_propia = 'mifuente';
$documento->addFontStyle(
    $fuente_propia,
    array(
        'name' => 'Arial',
        'size' => '14',
        'bold' => 'true',
        'color' => '5882FA'
    )
);
$seccion->addText(
    htmlspecialchars('Tercer texto con formato'),
    $fuente_propia
);

// Texto con objetos
$fuente = new Font();
$fuente->setBold(true);
$fuente->setName('Tahoma');
$fuente->setSize(16);
$fuente->setColor('9F81F7');
$texto = $seccion->addText(htmlspecialchars(
    'Cuarto texto con formato'
));
$texto->setFontStyle($fuente);

// Tabla personalizada
$estilo_tabla = array(
    'borderColor' => 'F2F2F2',
    'borderSize' => '5',
    'cellMargin' => '20',
    'bgColor' => '088A68',
);

$primera_fila = array('bgColor' => 'F2F2F2');
$documento->addTableStyle('mitabla',$estilo_tabla, $primera_fila);
$tabla = $seccion->addTable('mitabla');
for ($row = 1; $row <= 8; $row++) {
    $tabla->addRow();
    for ($cell = 1; $cell <= 3; $cell++) {
        if ($row == 1)
            $tabla->addCell(200)->addText(htmlspecialchars('primera'));
        else
            $tabla->addCell(200)->addText(htmlspecialchars('celda'));
    }
}

// Espacio
$seccion->addTextBreak(1);

// Imagen
$seccion->addImage(
		'imagen.jpg',
		array(
				'width' => 450,
				'height' => 300,
				'wrappingStyle' => 'behind'
		)
);

// Guardando documento
// $objWriter = IOFactory::createWriter($documento, "Word2007");
$documento->save("Documento01.docx", "Word2007");

header("Content-Disposition: attachment; filename=Documento01.docx");
echo file_get_contents('Documento01.docx');
