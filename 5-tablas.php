<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 5:
 * Tablas
 * Nota: se utiliza la notación corta de arreglos [], que se pueden remplazar por array(),
 * más información en https://parzibyte.me/blog/2018/10/11/sintaxis-corta-array-php/
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Language;

$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Tablas");

# Agregar texto...
/*
Todos los textos deben estar dentro de una sección
 */

$seccion = $documento->addSection();
$estiloTabla = [
    "borderColor" => "8bc34a",
    "alignment" => Jc::CENTER,
    "borderSize" => 5,
];
// Guardarlo para usarlo más tarde
$documento->addTableStyle("estilo1", $estiloTabla);

$tabla = $seccion->addTable("estilo1"); # Agregar tabla con el estilo que guardamos antes
$tabla->addRow(); # Agregar fila
$celda = $tabla->addCell(); # Agregar celda
$celda->addText("Dentro de una celda");
$celda = $tabla->addCell(); # Agregar celda
$celda->addText("Dentro de una celda");

# Un separador
$seccion->addText("Aquí otra tabla:");

# Otra tabla
$estiloTabla = [
    "borderColor" => "000000",
    "alignment" => Jc::LEFT,
    "borderSize" => 10,
    "cellMargin" => 10,
];
// Guardarlo para usarlo más tarde
$documento->addTableStyle("estilo2", $estiloTabla);
$tabla = $seccion->addTable("estilo2");
for ($fila = 0; $fila < 5; $fila++) {
    $tabla->addRow();
    for ($numeroCelda = 0; $numeroCelda < 5; $numeroCelda++) {
        $celda = $tabla->addCell();
        $celda->addText(sprintf("Posición %d x %d", $fila, $numeroCelda));
    }
}

# Otra tabla
$estiloTabla = [
    "borderColor" => "000000",
    "alignment" => Jc::RIGHT,
    "borderSize" => 30,
    "cellMargin" => 80,
];
// Guardarlo para usarlo más tarde
$documento->addTableStyle("estilo3", $estiloTabla);
$tabla = $seccion->addTable("estilo3");
$mascotas = [
    [
        "nombre" => "Maggie",
        "edad" => 3,
    ],
    [
        "nombre" => "Panqué",
        "edad" => 1,
    ],
    [
        "nombre" => "Guayaba",
        "edad" => 2,
    ],
];
# Encabezados
$fuente = [
    "name" => "Arial",
    "size" => 12,
    "color" => "000000",
];
$tabla->addRow();
$tabla->addCell()->addText("Nombre", $fuente);
$tabla->addCell()->addText("Edad", $fuente);

foreach ($mascotas as $mascota) {
    $tabla->addRow();
    $tabla->addCell()->addText($mascota["nombre"]);
    $tabla->addCell()->addText($mascota["edad"]);
}

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("5-tablas.docx");
