# Validación con el ZIP real

Prueba local realizada el 12 de octubre de 2026, con el R34 2026-10 y el parque de septiembre tal como fueron entregados. El parque incompatible no se trató como parque de octubre.

| Control | Cantidad |
|---|---:|
| Registros leídos | 634.601 |
| Registros RNC | 634.586 |
| Excluidos (CARROPAGO) | 15 |
| Total RESUMEN | 634.586 |
| COBRO CENTRALIZADO | 77.890 |
| PROFORMA | 346.908 |
| EXONERADO | 12.351 |
| PENDIENTE DE VALIDACION | 197.437 |
| Pendientes por pertenencia no reconocida | 195.379 |
| Pendientes por parque de período incompatible | 2.058 |
| Repeticiones de banco + AFIPOS en origen | 0 |
| Agentes / pertenencias conservadas | 67 |
| Bancos | 18 |
| Grupos agente/banco/modalidad/rango | 1.419 |

Leídos = RNC + excluidos y resumen = suma de modalidades = RNC. Todas las diferencias son cero. Las 67 agrupaciones incluyen pertenencias desconocidas identificadas como pendientes, no se afirma que todas sean agentes comerciales ya validados.

La primera ejecución completa tardó aproximadamente 49 segundos en este equipo, con bloques de 25.000 registros. Este tiempo no es una garantía para otros servidores.

## Verificaciones realizadas

- Pruebas automatizadas de las tres excepciones del parque, cuatro exonerados, exclusiones exactas y variantes del diccionario.
- Pruebas de duplicados entre bloques, referencias repetidas, equipos pequeños, conservación de identificadores de 20 dígitos y ceros iniciales, período inválido y reinicio entre procesamientos.
- Relectura independiente del Excel completo: cada AFIPOS, TERMINAL y SERIAL se compara con el CSV original, respetando el orden y las exclusiones. Las 1.419 agrupaciones del Excel coinciden con el resumen; se comprueba que ningún exonerado transa ni pertenece a otro agente.
- El archivo se abre con openpyxl y se inspecciona su estructura de dos hojas, autofiltros y cuatro listas de selección.
- Revisión visual de RNC, jerarquía RESUMEN y controles, mediante una copia de inspección que conserva el resumen completo y acota únicamente el detalle RNC para renderizar.
- Recálculo en Artifact Tool: total inicial 634.586; filtro EXONERADO 12.351; EXONERADO + TRANSANDO cero; VEPAGOS + BANCO DE VENEZUELA 17.011, todos centralizados.
- Streamlit AppTest: cuatro cargas, botón deshabilitado sin entradas, previsualización, descarga, filtro de modalidad y descarte de resultados al cambiar entradas.

No se abrió el archivo en Microsoft Excel de escritorio: se verificó con lectores y un motor de cálculo alternativo. No se afirma que se haya realizado una prueba de interfaz nativa de Excel.

## Información pendiente de los datos

Se necesita el parque de octubre para resolver los 2.058 casos especiales. El diccionario no resuelve BIOPAGO, SERVICIO VIRTUAL y varias pertenencias bancarias; esos equipos se conservan pendientes. Incorporar equivalencias o autorizar exclusiones adicionales exige información del negocio. No se inventaron precios ni se trasladaron resultados históricos de la plantilla.

## Ampliación: R34 en ZIP

La suite ampliada pasa 25 pruebas, incluyendo ZIP con CSV anidado, otros archivos ignorados, equivalencia de resultados con CSV directo, lectura incremental, ZIP ambiguo/dañado/con contraseña, límite descomprimido y encabezados UTF-8. Se procesó directamente el `ibo.zip` original: 634.601 registros leídos y 634.586 registros RNC. Los 1.419 grupos del resumen coinciden con los obtenidos del CSV directo.

El límite de subida de Streamlit continúa en 1.024 MB por archivo y ahora se muestra su valor activo en la interfaz. El CSV dentro del ZIP admite hasta 4 GB descomprimidos, sujeto al límite de filas de Excel y a los recursos disponibles del servidor.
