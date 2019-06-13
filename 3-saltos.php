<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 3:
 * Saltos de párrafo y saltos de página
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\Style\Language;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Saltos");

# Agregar texto...
/*
Todos los textos deben estar dentro de una sección
 */

$seccion = $documento->addSection();
# Simple texto
$seccion->addText("Hola, esto es algo de texto");
# Agregar 5 saltos
$seccion->addTextBreak(5);
$seccion->addText("Texto 5 saltos después");
# Agregar un salto de página
$seccion->addPageBreak();
$seccion->addText("Estoy en una nueva línea");
$seccion->addPageBreak();
$seccion->addText("Estoy en una nueva línea");

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("3-saltos.docx");
