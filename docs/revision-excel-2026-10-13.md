# Revisión del RNC Ingresos — 13 de octubre de 2026

La descarga tiene cantidades coherentes y las mismas 12 columnas del modelo. Se verificaron individualmente los 634.586 registros, sin diferencias de clasificación respecto de las reglas y referencias disponibles. Faltaba una tabla dinámica nativa: se incorporó sin sustituir el resumen existente ni cambiar las reglas comerciales.

## Archivos y alcance

- Modelo: RNC ingreso.xlsx del ZIP `C:\Users\victo\Desktop\ibo.zip`.
- Descarga recibida, conservada intacta: `C:\Users\victo\Downloads\RNC Ingresos Actualizado.xlsx`.
- Nueva salida: `outputs/revision-2026-10-13/RNC Ingresos Actualizado.xlsx`.
- Referencias: RNC342280.csv, DICCIONARIO NUEVO.xlsx, AGREGADOS EN COBRO.xlsx, Parque parcial Septiembre.xlsx y el Word del ZIP.

Se respetó la definición posterior del usuario: cuatro cargas manuales, todos los registros del R34 sin filtro obligatorio por mes y RNC generado automáticamente. El ejemplo solo se utilizó para estudiar la estructura; no se copiaron sus registros, cantidades ni cachés. No se exigió igualdad de cantidades entre R34 diferentes.

## Comparación con el ejemplo

Las doce columnas coinciden exactamente en nombre y orden: AFIPOS, TERMINAL, NOMBRE_BANCO, MARCA, MODELO, SERIAL, PERTENENCIA, PERTENENCIA2, ESTADO_POSV2, RANGO, MANEJADOR y FACTURACION.

El modelo agrupa PERTENENCIA2 → NOMBRE_BANCO, con subtotales por agente, total general y filtros FACTURACION y RANGO. La tabla nueva conserva esa jerarquía mediante AGENTE → BANCO y los mismos dos filtros. Incluye los 67 grupos de pertenencia/agente, incluso los no reconocidos y pequeños, y 18 bancos distintos.

La presentación no es idéntica: el modelo tiene tres hojas y su resumen está en Hoja1; la salida tiene RNC y RESUMEN, como se solicitó. Se conservan las columnas por modalidad, controles y selectores existentes. Se añadió la tabla dinámica al final de RESUMEN, accesible desde D6; en esta salida ocupa A1437:B1813.

El ejemplo estaba guardado con FACTURACION = EXONERADO y ocultaba CREDICARDPOS y #N/A. La salida comienza con todos los grupos para no omitir agentes aplicables. No se trasladaron filtros históricos como nuevas exclusiones. Su total histórico visible, 10.103, no es una cifra que el nuevo R34 deba reproducir.

## Cantidades verificadas

| Concepto | Registros |
|---|---:|
| Leídos del R34 | 634.601 |
| Excluidos: CARROPAGO | 15 |
| Conservados en RNC | 634.586 |
| Cobro Centralizado | 77.890 |
| Proforma | 346.908 |
| Exonerado | 12.351 |
| Pendiente de validación | 197.437 |
| Clasificación comercial resuelta | 437.149 |
| Total general, incluidos pendientes separados | 634.586 |

Se cumple 634.601 = 634.586 + 15. Las cuatro modalidades suman 634.586. Los pendientes se conservan separados y no representan autorización de cobro.

Los 1.419 grupos agente/banco/modalidad/rango coinciden con RNC. Cuadran todos los subtotales del resumen existente. Microsoft Excel confirmó también los 375 subtotales de agente y banco de la tabla nativa y los cuatro totales por modalidad, sin diferencias.

No hay repeticiones de banco + AFIPOS ni multiplicación de filas por los cruces en esta salida. Cada fila corresponde al R34, salvo las 15 exclusiones; ningún agente excluido aparece en RNC o en los totales. La comparación de las doce columnas de todas las filas entre la descarga recibida y la nueva salida no encontró cambios de valores.

## Reglas comerciales y pendientes

La auditoría independiente no utilizó la función clasificadora de la aplicación y no encontró discrepancias respecto de las reglas y referencias disponibles.

- Cobro Centralizado se cruza por agente y banco, sujeto a las excepciones del parque.
- Proforma se aplica según las condiciones del cruce; no se asigna automáticamente a pertenencias desconocidas.
- Exonerado exige agente autorizado, cero transacciones mensuales y rango compatible. Los 12.351 casos son DISGLOBAL 9.053, POSGLOBAL 3.075, UBIIPAGOS 222 y LATIN PAGOS 2021, 1. No se exoneraron casos TRANSANDO.
- Se mantienen las exclusiones CARROPAGO, POSITIVOPLUS, ACIERTOTECNOLOGICO y AGILPAGOS y sus equivalencias reconocidas. En este R34 solo CARROPAGO aporta excluidos.

RAPIDPAGO/BDV tiene 1.148 casos especiales; RAPIDPAGO/BANFANB, 315; SERVICIOS J&G/BDV, 595: total 2.058. El R34 corresponde a octubre de 2026 y el parque disponible a septiembre. Se conservaron correctamente como pendientes; falta el parque de octubre para determinar su modalidad definitiva. Las pruebas automatizadas cubren coincidencias, ausencias y conflictos con parques compatibles.

Otros 195.379 registros tienen pertenencias no resueltas por el diccionario, entre ellas BIOPAGO y SERVICIO VIRTUAL. No se inventaron equivalencias. Ambas causas suman los 197.437 pendientes.

También se reportan 74.133 discrepancias de POS_CON_TRANSACCION frente a los indicadores mensuales. Se conservó su tratamiento: se utilizan el conteo mensual y RANGO. Falta aclarar el alcance temporal de POS_CON_TRANSACCION si el negocio pretende darle otra prioridad.

## Tabla dinámica: funcionalidad y diferencias

La salida nueva contiene una tabla dinámica nativa reconocida por Microsoft Excel, con filtros de facturación y rango, selección múltiple habilitada, campos reorganizables, subtotales y total general. La caché se crea exclusivamente con los nuevos grupos de cada ejecución.

Su origen es el detalle agrupado H:L y calcula SUMA(EQUIPOS); el modelo usa CUENTA(AFIPOS). Los conteos son equivalentes para esta salida y se reconciliaron con cada fila del RNC. Al hacer doble clic devuelve grupos, no equipos individuales. Actualizar relee H:L: editar RNC manualmente no reclasifica ni reconstruye ese detalle. Los datos nuevos se procesan desde la aplicación.

Los selectores B1:B4, el autofiltro H:L y los filtros nativos funcionan independientemente. La vista de fórmulas conservada no reordena ni suprime filas de valor cero al filtrar. No se afirma equivalencia funcional absoluta ni identidad visual con el original.

## Cambios y comprobaciones

Se cambió únicamente la exportación, se añadió el módulo de tabla dinámica, pruebas y una explicación breve en Streamlit. Se conservaron las cuatro cargas, R34 CSV/ZIP, límite de 1.024 MB, procesamiento completo sin selección de mes y reglas comerciales. No se reconstruyó la aplicación.

Pasan 30 pruebas, incluyendo el botón real de Streamlit, cruces comerciales, CSV/ZIP, reprocesamiento y caché sin históricos. Se comprueba también la serialización XML compatible del paquete.

El archivo completo se abrió en Microsoft Excel 16.0 en modo normal y solo lectura, sin solicitar reparación. Excel reconoció dos hojas, 634.587 filas de RNC contando encabezados y una tabla dinámica. Actualizar mantuvo 634.586; aumentar temporalmente un grupo fuente en uno produjo 634.587. Se restauró el valor y se cerró sin guardar. Exonerado devuelve 12.351 tanto en fórmulas como en la tabla nativa. Los archivos originales no se sobrescribieron.

Los totales y la integridad de la exportación están comprobados. No se encontraron errores numéricos. Quedó resuelta la ausencia de tabla dinámica, con las diferencias de detalle descritas. Para resolver todas las clasificaciones faltan el parque de octubre y equivalencias válidas para las pertenencias desconocidas; no es correcto considerar los 197.437 pendientes listos para facturar. La comprobación cubre Excel de escritorio en esta máquina, no todas sus versiones ni el despliegue remoto de Streamlit.
