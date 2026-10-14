# Corrección definitiva y pruebas — 13 de octubre de 2026

## Qué se corrigió

La especificación nueva sustituye el alcance anterior que conservaba todo el R34. El código anterior incluía SITCA y CREDICARDPOS porque no estaban excluidos y retenía desconocidos como pendientes. Ahora solo exporta el alcance comercial: agentes del resumen del nuevo modelo, agentes reconocidos del archivo de cobro y excepciones de exoneración. Las equivalencias del diccionario se aplican antes de excluir. El diccionario por sí solo no habilita pertenencias ajenas.

El resumen anterior combinaba fórmulas, detalle agrupado, validaciones y una tabla dinámica adicional; por eso era grande. La nueva exportación tiene únicamente Hoja1 y RNC. Hoja1 contiene TablaDinámica1 en A:B, FACTURACION en A2 y PROFORMA en B2, encabezados en A4:B4 y un conteo por agente, sin bancos desplegados por defecto. La vista inicial de esta prueba termina en la fila 34.

El archivo recibido RNC Ingresos Actualizado (1).xlsx contiene además Hoja2 y Hoja3: son cuatro hojas en ese archivo concreto. El exportador previo del repositorio solo creaba RNC y RESUMEN; no hay evidencia para atribuirle esas dos hojas adicionales. La nueva salida se verificó con exactamente dos hojas.

Se incorporó CODIGO_AFILIADO y se adoptó el orden exacto de las 13 columnas del nuevo RNC.xlsx. El modelo anterior tenía 12, de ahí la diferencia. Se quitaron freeze_panes del exportador; no hay paneles congelados, divisiones ni protección de hojas. Se conserva el autofiltro normal de RNC, que no inmoviliza celdas.

El parque se cruza por AFIPOS completo normalizado, agente y banco. Una coincidencia válida ya no se rechaza por NUMPOS distinto ni por el nombre del mes del archivo. Si no hay coincidencia: PROFORMA. Solo cuando AFIPOS no es válido se intenta recuperar un identificador único mediante agente+banco+afiliado+terminal; no se acepta una coincidencia parcial ni ambigua. El usuario debe cargar el parque correspondiente al procesamiento.

POSGLOBAL da prioridad al cruce de cobro tanto TRANSANDO como SIN TX. Fuera de cobro, SIN TX se exonera y TRANSANDO es PROFORMA. Para DISGLOBAL, UBIIPAGOS y LATINPAGOS2021 sigue la exoneración cuando no transan. RANGO determina el estado, con valores explícitos; las contradicciones del conteo mensual se reportan sin confundir cualquier texto TX con ausencia de transacciones.

## Fuentes y límites de la comparación

Se inspeccionaron los dos Excel nuevos del usuario. RNC.xlsx tiene 636.210 registros históricos y la tabla compacta de referencia. La prueba completa se ejecutó con RNC342280.csv del ZIP original, de 634.601 registros, y las tres referencias operativas del mismo ZIP. No se localizó el R34 más reciente correspondiente a los 636.210 registros: no se reconstruyó ni inventó a partir del Excel de ejemplo. Por eso las cantidades no se forzaron a coincidir.

La lista del modelo se usa como configuración de alcance, sin sus cantidades ni registros; se amplía con agentes válidos de cobro y exoneración. Se conservan 33 agentes comerciales en esta prueba, incluidos CHIN CHIN, MEMORYTECH, SERVICIO TECNICO POS 2021 y LATIN PAGOS 2021 aunque no figuren en la vista inicial PROFORMA del ejemplo. No se exige cargar RNC.xlsx en Streamlit.

## Resultado de la prueba completa

| Concepto | Registros |
|---|---:|
| R34 leído | 634.601 |
| Excluidos | 484.872 |
| RNC comercial | 149.729 |
| Cobro Centralizado | 78.701 |
| Proforma | 59.465 |
| Exonerado | 11.563 |
| Pendientes | 0 |

634.601 = 484.872 + 149.729. Las modalidades suman 149.729. La vista inicial PROFORMA muestra 59.465; al quitar el filtro muestra 149.729. No debe confundirse el total filtrado con el total de RNC.

Se compararon las trece columnas de cada fila exportada contra el R34, diccionario, cobro y parque mediante una auditoría independiente de la función clasificadora: cero diferencias. No se detectaron repeticiones de banco+AFIPOS en el R34 de prueba. Las exclusiones no aparecen en RNC ni en la caché.

## Casos especiales

| Agente / banco | Centralizado | Proforma |
|---|---:|---:|
| RAPIDPAGO / Banco de Venezuela | 11 | 1.137 |
| RAPIDPAGO / Banfanb | 1 | 314 |
| SERVICIOS J&G / Banco de Venezuela | 11 | 584 |

Los 23 registros especiales del parque se localizaron individualmente. Estas cantidades son resultados de los archivos, no constantes del programa. RapidPago también tiene otros bancos: su total PROFORMA en la vista inicial es 3.044. El ejemplo nuevo muestra 3.048 y procede de otro R34; no se ajustaron cantidades artificialmente.

POSGLOBAL quedó con 2.861 centralizados, 2.287 exonerados y 1.926 Proforma. Los centralizados corresponden a Banco de Venezuela 2.468, Banco Exterior 233, Banco Plaza 96 y Mi Banco 64. Las pruebas cubren explícitamente las cuatro combinaciones solicitadas.

## Tabla dinámica y Microsoft Excel

TablaDinámica1 es nativa y su origen directo es RNC!A1:M149730. Tiene los trece campos y una caché nueva de equipos individuales. Ya no usa SUMA(EQUIPOS) sobre una tabla auxiliar: utiliza CUENTA(AFIPOS). No hay categorías heredadas del archivo de referencia.

Se abrió en Microsoft Excel 16.0 en modo normal, sin solicitar reparación. Se verificaron el filtro inicial, las tres modalidades presentes, total sin filtros, actualización y reorganización para añadir NOMBRE_BANCO. Se contrastaron los 256 grupos agente/banco/modalidad contra la auditoría independiente, sin diferencias. La reorganización requirió reemplazar la disposición de campos existente en la llamada de prueba; el intento de añadir otra vez el mismo campo fue rechazado por Excel.

La caché contiene datos individuales y Excel tiene habilitado mostrar detalle. La prueba automatizada de ShowDetail no creó una hoja de detalle; por tanto no se afirma haber verificado el gesto de doble clic en la interfaz. Sí se verificaron apertura, filtros, actualización y análisis por banco. La vista compacta y RNC fueron renderizadas y revisadas visualmente.

Se comprobó en el XML de ambas hojas la ausencia de elementos pane y sheetProtection. Las celdas no están protegidas. Se mantienen solo dos hojas al entregar; las modificaciones de comprobación en Excel se cerraron sin guardar.

Si no existe PROFORMA en una ejecución, se abre en (Todas), sin inventar categorías. Si todos los registros quedan excluidos, se entrega el esquema vacío con total cero, sin crear una tabla dinámica inválida.

## Código y pruebas

Se modificaron normalizacion.py (alcance y columnas), entradas.py (claves de parque), facturacion.py (reglas), procesamiento.py (retirar el bloqueo por mes), exportacion.py y pivot.py (salida compacta y caché individual). app.py conserva cargas, botones, vistas y descarga; solo actualiza los textos explicativos. Se actualizaron y añadieron pruebas y documentación. No se cambiaron el lector CSV/ZIP, límites de carga ni módulos de precios.

Pasan 34 pruebas automatizadas: cuatro cargas, botón real de Streamlit, CSV/ZIP, procesamiento por bloques, identificadores, duplicados, exclusiones, reglas POSGLOBAL, parque exacto y fallback ambiguo, rangos, caché nueva, trece columnas y ausencia de paneles/protección. Los datos operativos y el Excel de salida no se suben a GitHub.

## Exclusiones por pertenencia estandarizada

| Pertenencia | Registros |
|---|---:|
| SITCA | 157,923 |
| CREDICARDPOS | 131,555 |
| BIOPAGO | 99,836 |
| SERVICIO VIRTUAL | 54,529 |
| EXTERIOR | 15,886 |
| BANPLUS | 8,438 |
| BANCARIBE | 4,285 |
| BFC | 3,514 |
| BICENTENARIO | 2,954 |
| VENEZUELA | 1,963 |
| BANCAMIGA | 1,179 |
| ACTIVO | 1,072 |
| TESORO | 770 |
| MI BANCO | 412 |
| PLAZA | 192 |
| CARONI | 169 |
| MIBANCO | 71 |
| BANFANB | 40 |
| OFICINA PRINCIPAL CARACAS | 27 |
| CARROPAGO | 15 |
| BANCRECER | 13 |
| 100% BANCO | 6 |
| BANCO EXTERIOR | 4 |
| BCO VENEZUELA | 2 |
| OFICINA COMERCIAL PTO. LA CRUZ | 2 |
| OFICINA COMERCIAL NVA ESPARTA | 2 |
| OFICINA COMERCIAL BARQUISIMETO | 2 |
| OFICINA COMERCIAL MARACAY | 2 |
| OFICINA COMERCIAL VALENCIA | 2 |
| OFICINA COMERCIAL PTO. ORDAZ | 2 |
| MEGASOFT | 1 |
| BOD | 1 |
| OFICINA COMERCIAL MARACAIBO | 1 |
| OFICINA COMERCIAL SN CRISTOBAL | 1 |
| BANCO PLAZA | 1 |

El informe y los resultados corresponden a los archivos disponibles. Para certificar las cantidades del R34 más reciente hay que cargar ese R34 y sus referencias en la aplicación corregida.
