# Automatización de Agregados Comerciales

Aplicación Python y Streamlit para procesar el **R34 actual** y descargar **RNC Ingresos Actualizado.xlsx**. No acumula históricos ni lee archivos operativos del repositorio.

## Uso

Con Python 3.11 o 3.12:

```sh
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

1. Seleccione el mes de procesamiento.
2. Cargue **R34 (CSV)**, **DICCIONARIO NUEVO**, **AGREGADOS EN COBRO** y **PARQUE PARCIAL (XLSX)**.
3. Pulse **Procesar archivos** y revise RNC, RESUMEN y las validaciones.
4. Descargue el único Excel. Contiene exactamente **RNC** y **RESUMEN**. Las validaciones están dentro de RESUMEN.

Cambiar cualquiera de las entradas o el período invalida la descarga anterior. No se solicita la plantilla RNC ingreso.xlsx. Los archivos originales del ZIP son ejemplos para análisis y pruebas; no se distribuyen ni se usan en producción. El Word en `docs` conserva los requisitos originales como referencia histórica: **la definición del apartado 20 tiene prioridad**, incluida la carga manual de los cuatro archivos y las dos hojas de salida.

## Salida y filtros

**RNC** conserva las 12 columnas de la plantilla: AFIPOS, TERMINAL, NOMBRE_BANCO, MARCA, MODELO, SERIAL, PERTENENCIA, PERTENENCIA2, ESTADO_POSV2, RANGO, MANEJADOR y FACTURACION. Incluye autofiltro y encabezados inmovilizados. AFIPOS, TERMINAL y SERIAL se leen y escriben como texto, sin conversiones a flotantes.

**RESUMEN** reproduce la jerarquía de Hoja1: agente estandarizado → bancos, cuenta de equipos, subtotal de cada agente y total general. Agrega columnas por modalidad. Los selectores **B1:B4** filtran facturación, rango, agente y banco mediante fórmulas SUMIFS. Se inicia con todos los registros, sin ocultar agentes pequeños. Los agentes se ordenan por su total inicial, como en el modelo; cambiar filtros recalcula cantidades pero no reordena ni elimina las filas con cero.

Es un **resumen equivalente con fórmulas**, no una tabla dinámica nativa. Para selección múltiple se usan los filtros de Streamlit o el detalle agrupado **H:L** de Excel. Los autofiltros H:L tienen su propio subtotal visible L3 y son independientes de los selectores B1:B4. Las columnas **N:R** contienen controles, exclusiones, pertenencias desconocidas y hasta 300 ejemplos de incidencias. Los conteos de incidencias y de desconocidos son completos. Un registro puede tener varias incidencias.

Las fórmulas filtran el resumen calculado al procesar, no vuelven a clasificar manualmente un RNC editado en Excel. Para actualizar datos, cargue los cuatro archivos y procese de nuevo. Excel recibe valores calculados iniciales y recalcula al modificar selectores.

## Reglas y controles

- El diccionario se carga en cada ejecución. Se normalizan mayúsculas, espacios y acentos para buscar equivalencias. Mapeos ambiguos detienen el proceso; celdas vacías o errores `#N/A` no se consideran agentes válidos.
- Pertenencias desconocidas conservan el original y quedan como **PENDIENTE DE VALIDACION**, sin adivinar empresas ni aplicar PROFORMA por defecto.
- AGREGADOS EN COBRO se interpreta como **una columna jerárquica**, agente seguido de sus bancos. El cruce siempre usa **agente + banco**. Un agente no reconocido o un banco sin agente detienen la lectura.
- Se excluyen CARROPAGO, POSITIVOPLUS, ACIERTOTECNOLOGICO y AGILPAGOS, incluidos sus equivalentes exactos del diccionario. No se usan coincidencias parciales.
- DISGLOBAL, UBIIPAGOS, POSGLOBAL y LATINPAGOS2021 se exoneran solo cuando **CANT_TRANS_ACUM_MES = 0** y RANGO indica **SIN TX EN EL MES ACTUAL, 30, 60, 90 o 120 DIAS SIN TX**. TRANSANDO debe tener conteo positivo entero. Contradicciones, valores negativos o inválidos quedan pendientes.
- POS_CON_TRANSACCION se reporta cuando contradice los indicadores mensuales. No sustituye a estos: en el R34 real aparece como 1 en miles de registros con conteo mensual cero y rango sin transacciones; su alcance temporal no está documentado.
- RAPIDPAGO/BDV, RAPIDPAGO/BANFANB y SERVICIOS J&G/BDV se cruzan con `7A_PARQUE_PARCIAL_ACTIVOS` por **agente + banco + AFIPOS + NUMPOS/TERMINAL**. En R34, TERMINAL es una descripción de hardware, por eso se usa NUMPOS contra la columna TERMINAL numérica del parque.
- Un equipo especial ausente de un parque válido es PROFORMA. Si su AFIPOS existe pero discrepan otras claves, queda pendiente. Sin parque del período no se inventa una clasificación.
- Se reconocen 201 = BANCO DE VENEZUELA y 896 = BANFANB, incluidas variantes BANFAB/BANFAN. Códigos y nombres contradictorios quedan pendientes.
- El R34 debe corresponder por completo al año/mes seleccionado. El parque de ejemplo no tiene fecha interna: el usuario debe cargar el del mes. Un mes o año incompatible en su nombre deja los casos especiales pendientes y lo informa.
- Las referencias duplicadas se convierten en claves únicas; los cruces no multiplican registros. Se detectan repeticiones originales de banco + AFIPOS en disco y se conservan para revisión, porque no hay autorización de eliminación.
- Se valida **leídos = RNC + excluidos**, **resumen = RNC** y **suma de modalidades = RNC** antes de publicar el archivo. Los pendientes se cuentan por separado. No se excluyen otros agentes basándose en filtros antiguos de la plantilla.

## Archivos grandes y despliegue

El CSV se lee por bloques de 25.000 filas y solo con las columnas necesarias. El Excel se escribe secuencialmente en disco mediante XlsxWriter. La detección de claves repetidas usa SQLite temporal local, sin servicio externo. La previsualización muestra como máximo 1.000 registros.

Streamlit mantiene las subidas y descargas en memoria: un CSV de 626 MB necesita un servidor con margen suficiente. Se recomienda un equipo local o Codespaces con **al menos 4 GB de RAM y varios GB libres de disco temporal**. El límite de subida configurado es 1 GB por archivo; subirlo no incrementa los recursos del servidor. Excel admite como máximo 1.048.575 registros más encabezado en RNC; superar ese límite produce un error explícito.

Codespaces incluye configuración en `.devcontainer`; ejecute `streamlit run app.py` y abra el puerto 8501. Para Streamlit Community Cloud seleccione este repositorio y `app.py`, pero valide su capacidad con el tamaño real. Si el servidor no soporta la carga, ejecute localmente/Codespaces o use la alternativa por archivos en disco:

```sh
python tools/process_files.py --r34 R34.csv --diccionario "DICCIONARIO NUEVO.xlsx" --cobro "AGREGADOS EN COBRO.xlsx" --parque "Parque octubre.xlsx" --periodo 2026-10 --salida "RNC Ingresos Actualizado.xlsx"
```

No se registran ni envían archivos de clientes a GitHub. `.gitignore` excluye CSV, XLSX, ZIP y salidas temporales.

## Pruebas y documentación

```sh
python -m unittest discover -s tests -v
```

Incluye reglas comerciales, cruces compuestos, duplicados entre bloques, identificadores largos, filtros, cargas obligatorias, cambios de entrada, conservación de agentes pequeños y ausencia de acumulación entre procesamientos. GitHub Actions ejecuta las pruebas con Python 3.12.

Consulte `docs/analisis-plantilla.md` y `docs/validacion-real.md` para la estructura observada, decisiones y resultados de la prueba completa del ZIP.

## Precios e histórico

`config/precios.ejemplo.json` y `agregados/precios.py` preparan la configuración futura de agente, banco, modalidad, precio, moneda, condiciones y vigencia. No hay tarifas inventadas ni cálculo monetario. El procesamiento es una función independiente por período; cualquier modo histórico requiere definir claves, actualización y permanencia antes de implementarlo.
