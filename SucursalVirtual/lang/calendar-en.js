// ** I18N

// Calendar EN language
// Author: Mihai Bazon, <mihai_bazon@yahoo.com>
// Encoding: any
// Distributed under the same terms as the calendar itself.

// For translators: please use UTF-8 if possible.  We strongly believe that
// Unicode is the answer to a real internationalized world.  Also please
// include your contact information in the header, as can be seen above.

// full day names
Calendar._DN = new Array
("Domingo",
 "Lunes",
 "Martes",
 "Miércoles",
 "Jueves",
 "Viernes",
 "Sábado",
 "Domingo");


Calendar._SDN = new Array
("Dom",
 "Lun",
 "Mar",
 "Mié",
 "Jue",
 "Vie",
 "Sab",
 "Dom");


Calendar._FD = 0;

Calendar._MN = new Array
("Enero",
 "Febrero",
 "Marzo",
 "Abril",
 "Mayo",
 "Junio",
 "Julio",
 "Agosto",
 "Septiembre",
 "Octubre",
 "Noviembre",
 "Diciembre");

// short month names
Calendar._SMN = new Array
("Ene",
 "Feb",
 "Mar",
 "Abr",
 "May",
 "Jun",
 "Jul",
 "Ago",
 "Sep",
 "Oct",
 "Nov",
 "Dic");

// tooltips
Calendar._TT = {};
Calendar._TT["INFO"] = "Acerca";

Calendar._TT["ABOUT"] =
"::SELECCION DE FECHA::\n" +
"\n\n" +
"Afex.cl (c) 2005\n" + // don't translate this this ;-)
"Modo de Uso: Elija la fecha a designar para la entrega de la Operación Asignada.\n" +
"Puede navegar por los botones y seleccionar : dias, meses, años." +
"\n\n" +
"Date selection:\n" +
"- Use los " + String.fromCharCode(0x2039) + ", " + String.fromCharCode(0x203a) + " botones para seleccionar el mes.\n" +
"- Mantenga presionado el boton del Mouse y mueva la ventana del Calendario.";
Calendar._TT["ABOUT_TIME"] = "\n\n" +
"Time selection:\n" +
"- Click on any of the time parts to increase it\n" +
"- or Shift-click to decrease it\n" +
"- or click and drag for faster selection.";

Calendar._TT["PREV_YEAR"] = "Prev. Año ( )";
Calendar._TT["PREV_MONTH"] = "Prev. Mes ( )";
Calendar._TT["GO_TODAY"] = "Ir Hoy";
Calendar._TT["NEXT_MONTH"] = "Prox. Mes ( )";
Calendar._TT["NEXT_YEAR"] = "Prox.Año ( )";
Calendar._TT["SEL_DATE"] = "Seleccionar fecha";
Calendar._TT["DRAG_TO_MOVE"] = "Arrastrar y Mover";
Calendar._TT["PART_TODAY"] = " (Hoy)";


Calendar._TT["DAY_FIRST"] = "Mostrar %s primero";

Calendar._TT["WEEKEND"] = "0,6";

Calendar._TT["CLOSE"] = "Cerrar";
Calendar._TT["TODAY"] = "Hoy";
Calendar._TT["TIME_PART"] = "(Shift-)Click o mueva para cambiar el valor";


Calendar._TT["DEF_DATE_FORMAT"] = "%Y-%m-%d";
Calendar._TT["TT_DATE_FORMAT"] = "%a, %b %e";

Calendar._TT["WK"] = "Sem";
Calendar._TT["TIME"] = "Hora:";
