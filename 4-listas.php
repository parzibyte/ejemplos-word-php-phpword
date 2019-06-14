<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 4:
 * Listas
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\Style\Language;
use PhpOffice\PhpWord\Style\ListItem;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Listas");

$seccion = $documento->addSection();
$seccion->addListItem("Elemento con profundidad por defecto");
$seccion->addListItem("Elemento con profundidad 1", 1);
$seccion->addListItem("Elemento con profundidad 1", 1);
$seccion->addListItem("Elemento con profundidad 2", 2);
$seccion->addListItem("Elemento con profundidad 2", 2);
$seccion->addListItem("Elemento con profundidad 3", 3);
$seccion->addListItem("Elemento con profundidad 3", 3);
$fuente = [
    "name" => "Courier new",
    "size" => 20,
    "color" => "000000",
    "italic" => true,
];
$seccion->addListItem("Elemento con profundidad 1 y con fuente", 1, $fuente);
for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_SQUARE_FILLED", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_SQUARE_FILLED,
    ]);
}

for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_BULLET_FILLED", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_BULLET_FILLED,
    ]);
}

for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_BULLET_EMPTY", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_BULLET_EMPTY,
    ]);
}

for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_NUMBER", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_NUMBER,
    ]);
}

for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_NUMBER_NESTED", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_NUMBER_NESTED,
    ]);
}

for ($profundidad = 1; $profundidad < 4; $profundidad++) {
    $seccion->addListItem("Elemento con profundidad $profundidad, con fuente y tipo de lista TYPE_ALPHANUM", $profundidad, $fuente, [
        'listType' => ListItem::TYPE_ALPHANUM,
    ]);
}

# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("4-listas.docx");
