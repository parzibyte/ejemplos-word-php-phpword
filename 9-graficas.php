<?php
/**
 * Trabajar con documentos de Word y PHP usando PHPOffice
 *
 * Más tutoriales en: parzibyte.me/blog
 *
 * Ejemplo 9:
 * Gráficas
 * Nota: se utiliza la notación corta de arreglos [], que se pueden remplazar por array(),
 * más información en https://parzibyte.me/blog/2018/10/11/sintaxis-corta-array-php/
 */
require_once "vendor/autoload.php";
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Language;
$documento = new \PhpOffice\PhpWord\PhpWord();
$propiedades = $documento->getDocInfo();
$propiedades->setCreator("Luis Cabrera Benito");
$propiedades->setTitle("Gráficas");

$seccion = $documento->addSection();

# Títulos. Solo modificando depth (el número)
$fuenteTitulo = [
    "name" => "Verdana",
    "size" => 20,
    "color" => "000000",
];
$documento->addTitleStyle(1, $fuenteTitulo);
$estilo = [
    "width" => Converter::cmToEmu(17),
    "height" => Converter::cmToEmu(10),
];
# Gráfica con una sola línea
$seccion->addTitle("Ventas del año actual", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
# Gráfica con 3 líneas
$grafica = $seccion->addChart("line", $etiquetas, $series, $estilo);
$seccion->addTitle("Comparación de ventas", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
$grafica = $seccion->addChart("line", $etiquetas, $series, $estilo);
# Agregar más datos...
$grafica->addSeries($etiquetas, [500, 132, 32, 432, 332, 456, 212, 5333, 4568, 123, 879, 4544]);
$grafica->addSeries($etiquetas, [999, 4848, 4544, 7833, 4549, 454, 212, 666, 121, 999, 454, 335]);
# La misma de arriba, pero no de tipo line, sino bar
$seccion->addTitle("Comparación de ventas", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
$grafica = $seccion->addChart("bar", $etiquetas, $series, $estilo);
# Agregar más datos...
$grafica->addSeries($etiquetas, [500, 132, 32, 432, 332, 456, 212, 5333, 4568, 123, 879, 4544]);
$grafica->addSeries($etiquetas, [999, 4848, 4544, 7833, 4549, 454, 212, 666, 121, 999, 454, 335]);
# La misma de arriba, pero no de tipo bar, sino column
$seccion->addTitle("Comparación de ventas", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
$grafica = $seccion->addChart("column", $etiquetas, $series, $estilo);
# Agregar más datos...
$grafica->addSeries($etiquetas, [500, 132, 32, 432, 332, 456, 212, 5333, 4568, 123, 879, 4544]);
$grafica->addSeries($etiquetas, [999, 4848, 4544, 7833, 4549, 454, 212, 666, 121, 999, 454, 335]);
# La misma de arriba, pero no de tipo column, sino area
$seccion->addTitle("Comparación de ventas", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
$grafica = $seccion->addChart("area", $etiquetas, $series, $estilo);
# Agregar más datos...
$grafica->addSeries($etiquetas, [500, 132, 32, 432, 332, 456, 212, 5333, 4568, 123, 879, 4544]);
$grafica->addSeries($etiquetas, [999, 4848, 4544, 7833, 4549, 454, 212, 666, 121, 999, 454, 335]);

# Una de pastel
$seccion->addTitle("Gastos por categoría", 1);
$etiquetas = ["Comida", "Escuela", "Tecnología"];
$series = [123, 456, 789];
$grafica = $seccion->addChart("pie", $etiquetas, $series, $estilo);
# Una de dona
$seccion->addTitle("Gastos por categoría", 1);
$etiquetas = ["Comida", "Escuela", "Tecnología"];
$series = [123, 456, 789];
$grafica = $seccion->addChart("doughnut", $etiquetas, $series, $estilo);
# La primera, pero en 3d
$estilo = [
    "width" => Converter::cmToEmu(17),
    "height" => Converter::cmToEmu(10),
    "3d" => true,
];
$seccion->addTitle("Ventas del año actual", 1);
$etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
$series = [600, 700, 999, 1000, 2800, 3000, 5000, 3000, 6000, 6000, 5000, 7000];
$grafica = $seccion->addChart("line", $etiquetas, $series, $estilo);



# Para que no diga que se abre en modo de compatibilidad
$documento->getCompatibility()->setOoxmlVersion(15);
# Idioma español de México
$documento->getSettings()->setThemeFontLang(new Language("ES-MX"));

# Guardarlo
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($documento, "Word2007");

$objWriter->save("9-graficas.docx");