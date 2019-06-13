<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 8:
 * Agregar contenido y un índice
 * Nota: se utiliza la notación corta de arreglos [], que se pueden remplazar por array(),
 * más información en https://parzibyte.me/blog/2018/10/11/sintaxis-corta-array-php/
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\Style\Language;
use PhpOffice\PhpWord\Style\TOC;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Tabla de contenidos");

/*
Todos los textos deben estar dentro de una sección
 */

$seccion = $documento->addSection();
# Registrar el estilo del título
$fuenteTitulo = [
    "name" => "Verdana",
    "size" => 20,
    "color" => "000000",
];
$documento->addTitleStyle(1, $fuenteTitulo);
# Agregar el título del índice
$seccion->addTitle("Índice", 1);
# Aquí agregamos la tabla de contenidos
$fuenteTablaContenidos = [
    "name" => "Arial",
    "size" => 20,
    "color" => "881111",
];
$estiloTablaDeContenidos = [
    "tabLeader" => TOC::TABLEADER_UNDERSCORE,
];
$seccion->addTOC($fuenteTablaContenidos, $estiloTablaDeContenidos);

$seccion->addTitle("Lorem", 1);
# Texto bajo el título
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addTitle("Ipsum", 1);
# Texto bajo el título
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addTitle("Dolor", 1);
# Texto bajo el título
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
# Ahora un subtítulo con profundidad de 2
$fuenteSubtitulo = [
    "name" => "Verdana",
    "size" => 18,
    "color" => "000000",
];
$documento->addTitleStyle(2, $fuenteSubtitulo);
$seccion->addTitle("Soy un subtítulo", 2);
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");

$seccion->addTitle("Soy un subtítulo", 2);
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");
$seccion->addTitle("Otro título", 1);
$seccion->addTitle("Otro subtítulo", 2);
# Texto bajo el título
$seccion->addText("Lorem ipsum dolor sit amet consectetur adipiscing elit, cursus facilisi id risus aliquet enim, varius ultricies dictum suspendisse mollis non.");

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("8-tabla-de-contenidos.docx");
