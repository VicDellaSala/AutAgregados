# Análisis del modelo y de los archivos

Se inspeccionaron el Word completo, el encabezado y todos los indicadores mensuales del CSV, las dos referencias pequeñas, el parque y la estructura XML de la tabla dinámica de RNC ingreso.xlsx. Los datos de la plantilla no forman parte de la aplicación.

## Hoja1 de la plantilla

- A1: filtro FACTURACION, guardado con EXONERADO seleccionado.
- A2: filtro RANGO, guardado con todas las categorías.
- A4:B40: jerarquía compacta, «Etiquetas de fila» y «Cuenta de AFIPOS».
- Campos de fila: PERTENENCIA2 primero y NOMBRE_BANCO después.
- Agentes ordenados por cantidad descendente; bancos anidados, subtotales por agente y total general.
- El archivo guarda agentes ocultos (incluido CREDICARDPOS) y una categoría #N/A. Son filtros de un informe anterior, no reglas nuevas de exclusión.
- Las cifras y anotaciones sueltas a la derecha no tienen una regla de cálculo documentada y no se trasladan como cantidades actuales ni tarifas.

La salida mantiene la jerarquía, cuenta, subtotales y filtros de facturación/rango; agrega filtros de agente/banco y columnas de modalidad. Se inicia en «Todas» para conservar todos los agentes aplicables. No se copia el filtro inicial EXONERADO porque ocultaría modalidades y agentes del procesamiento completo solicitado.

Hay una inconsistencia interna en la referencia: la caché de la tabla dinámica conserva **15 campos**, mientras su origen declara `RNC!A1:L629675`, de **12 columnas**. Hoja2 tiene 15 columnas y 523 registros de ejemplo; RNC tiene 12 columnas y 629.674 registros. Se utiliza la estructura de 12 columnas de RNC y se reconstruye el resumen sin importar esa caché antigua.

La salida es un resumen equivalente con SUMIFS y selectores, no una tabla dinámica nativa. El detalle agrupado permite autofiltros y subtotal visible sin duplicar filas ni subtotales en su base. Las cantidades se recalculan exclusivamente del R34 cargado.

## Entradas reales

- R34: CSV UTF-8 separado por punto y coma, 626.616.455 bytes, 634.601 registros, período 2026-10.
- Diccionario: encabezados RNC34/CORRECTO en B1:C1; la primera columna está vacía. Tiene alias, nombres ya normalizados y 10 destinos #N/A que no son estandarizaciones válidas.
- Cobro: una columna, 136 filas; los agentes están en negrita y los bancos con sangría. Tres combinaciones especiales están resaltadas en amarillo. Se conservan las asociaciones jerárquicas.
- Parque: hoja 7A_PARQUE_PARCIAL_ACTIVOS, 37 filas y 7 columnas. No contiene fecha del período dentro de la tabla. El nombre indica septiembre, incompatible con el R34 de octubre.
- R34.TERMINAL es una descripción de hardware; parque.TERMINAL es un número de terminal. La clave correspondiente es R34.NUMPOS. AFIPOS se valida junto con agente, banco y ese número.

## Transacciones

El R34 contiene seis rangos. Todos los TRANSANDO tienen CANT_TRANS_ACUM_MES > 0; los cinco rangos sin transacciones tienen conteo mensual cero. Se exigen ambas señales concordantes para exonerar. No se utilizan columnas de meses anteriores para determinar actividad del mes actual.

POS_CON_TRANSACCION = 1 aparece en **74.133** registros con conteo mensual cero y rango sin transacciones. No tiene una definición temporal proporcionada; utilizarlo como veto de exoneración contradiría los dos campos mensuales. Se informa como incidencia sin sustituir la condición documentada de conteo mensual + RANGO. Una contradicción entre estos últimos sí impide clasificar automáticamente.

## Decisiones conservadoras

Las pertenencias no estandarizables, claves incoherentes y excepciones sin parque válido permanecen en RNC, con modalidad PENDIENTE DE VALIDACION. Se cuentan separadamente y no se asignan a empresas conocidas por parecido. Los únicos eliminados son los cuatro agentes autorizados y sus equivalencias exactas.

El modo vigente utiliza exclusivamente R34 actual. El Word se conserva como antecedente; las instrucciones de carga interna, histórico o tercera hoja quedan sustituidas por el apartado 20 del usuario.
