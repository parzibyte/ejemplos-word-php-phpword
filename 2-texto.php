<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 2:
 * Agregar enlaces y texto con distintas fuentes y colores
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\Style\Language;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Texto");

# Agregar texto...
/*
Todos los textos deben estar dentro de una sección
 */

$seccion = $documento->addSection();
# Simple texto
$seccion->addText("Hola, esto es algo de texto");
# Con fuentes personalizadas
$fuente = [
    "name" => "Arial",
    "size" => 12,
    "color" => "8bc34a",
    "italic" => true,
    "bold" => true,
];
$seccion->addText("Hola, esto es algo de texto", $fuente);
# Hipervínculo
$fuenteHipervinculo = [
    "name" => "Arial",
    "size" => 12,
    "color" => "ff0000",
    "italic" => true,
];
$seccion->addLink("https://parzibyte.me/blog", "Mi blog", $fuenteHipervinculo);

# Títulos. Solo modificando depth (el número)
$fuenteTitulo = [
    "name" => "Verdana",
    "size" => 20,
    "color" => "000000",
];
$documento->addTitleStyle(1, $fuenteTitulo);
$seccion->addTitle("Soy un título", 1);
# Texto bajo el título
$seccion->addText("Hola");
# Ahora un subtítulo con profundidad de 2
$fuenteSubtitulo = [
    "name" => "Verdana",
    "size" => 18,
    "color" => "000000",
];
$documento->addTitleStyle(2, $fuenteSubtitulo);
$seccion->addTitle("Soy un subtítulo", 2);

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("2-texto.docx");
