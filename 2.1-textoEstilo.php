<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 2.1:
 * Estilizar el texto con textRun
 * Nota: se utiliza la notación corta de arreglos [], que se pueden remplazar por array(),
 * más información en https://parzibyte.me/blog/2018/10/11/sintaxis-corta-array-php/
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Language;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Texto con estilos");

# Agregar texto...
/*
Todos los textos deben estar dentro de una sección
 */

$seccion = $documento->addSection();
# Títulos. Solo modificando depth (el número)
$fuenteTitulo = [
    "name" => "Verdana",
    "size" => 20,
    "color" => "000000",
];
$documento->addTitleStyle(1, $fuenteTitulo);
$seccion->addTitle("Cotizaciones web", 1);
$textRun = $seccion->addTextRun([
    "alignment" => Jc::BOTH,
    "lineHeight" => 1, # Quedará muy pegado
]);
$fuente = [
    "name" => "Arial",
    "size" => 12,
    "color" => "8bc34a",
    "italic" => true,
    "bold" => true,
];

$textRun->addText("Un sistema web con PHP y MySQL que permite crear clientes y a partir de ellos cotizaciones con el costo automático, así como el tiempo de la cotización. ");
$textRun->addTextBreak(2);
$textRun->addText("Más tarde, eso se puede imprimir. Aparte de eso, se cuenta con el apartado de los ajustes, en donde se personalizan algunos mensajes");
$textRun->addTextBreak(2);
$textRun->addText("Hice el sistema porque personalmente necesitaba un software para cotizaciones que a veces son requeridas por mis clientes");

$textRun->addTextBreak(5);

$textRun->addText("Texto con una fuente, en el mismo párrafo ", $fuente);

$fuente = [
    "name" => "Verdana",
    "size" => 10,
    "color" => "00ff00",
];
$textRun->addText("que continúa aquí con otra fuente ", $fuente);
$fuente = [
    "name" => "Courier new",
    "size" => 8,
    "color" => "0000ee",
];
$textRun->addText("y sigue por aquí...", $fuente);

# Se pueden agregar más textruns, este va alineado al centro

$otroTextRun = $seccion->addTextRun([
    "alignment" => Jc::CENTER,
    "lineHeight" => 0.7,
]);
$fuente = [
    "name" => "Century Gothic",
    "size" => 15,
    "color" => "000000",
];

$otroTextRun->addText("Lorem ipsum dolor sit amet consectetur adipiscing, elit nulla et aptent ultricies inceptos, tristique torquent lacinia auctor integer. Facilisi eu tempus donec platea inceptos diam dis aliquam mi, vitae senectus ullamcorper nisi torquent auctor vehicula. Viverra rhoncus vestibulum ante bibendum dui volutpat duis auctor dictumst nulla, risus feugiat fusce nisl semper urna nullam aliquam.", $fuente);
$otroTextRun->addTextBreak(2);
$otroTextRun->addText("Pharetra pulvinar curae ac ante risus vestibulum mus diam neque, facilisi scelerisque dignissim velit suscipit ultrices nostra. Laoreet vivamus sem pretium nisi risus natoque magnis cubilia, aliquet eleifend posuere imperdiet dictum sociosqu vel fringilla, diam luctus penatibus eu at ultricies praesent. Nulla habitasse duis felis nostra senectus dapibus, sociosqu porttitor interdum scelerisque tortor donec pharetra, enim ligula dignissim hac nisl.", $fuente);
$otroTextRun->addTextBreak(2);

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("2.1-textoEstilo.docx");
