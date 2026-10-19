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

1. Cargue **R34 (CSV o ZIP)**, **DICCIONARIO NUEVO**, **AGREGADOS EN COBRO** y **PARQUE PARCIAL (XLSX)**.
2. Pulse **Procesar archivos** y revise RNC, RESUMEN y las validaciones.
3. Descargue el único Excel. Contiene exactamente **Hoja1** y **RNC**. Las validaciones se consultan en Streamlit.

Cambiar cualquiera de las entradas invalida la descarga anterior. No se selecciona mes: se leen todos los registros del R34 y se exporta el alcance comercial, aunque incluya varios períodos o no tenga columnas de fecha. La codificación queda en opciones de lectura plegadas. No se solicita la plantilla RNC ingreso.xlsx. Los archivos originales del ZIP son ejemplos para análisis y pruebas; no se distribuyen ni se usan en producción. El Word en `docs` conserva los requisitos originales como referencia histórica: **la definición del apartado 20 tiene prioridad**, incluida la carga manual de los cuatro archivos y las dos hojas de salida.

## Salida y reglas vigentes

La corrección definitiva del usuario sustituye el alcance anterior. Se lee el R34 completo, sin selección de mes, pero se exportan **solo los agentes comerciales aplicables**. El alcance reúne los agentes de Hoja1 del modelo RNC.xlsx, los agentes reconocidos de AGREGADOS EN COBRO y las excepciones de exoneración. El diccionario estandariza las variantes; una entrada del diccionario por sí sola no habilita un agente ajeno al alcance.

Se excluyen CREDICARDPOS y sus equivalencias, CARROPAGO, POSITIVOPLUS, ACIERTOTECNOLOGICO, AGILPAGOS y las pertenencias fuera de alcance. Los conteos por pertenencia se muestran en Validaciones de Streamlit.

El Excel tiene exactamente dos hojas, sin paneles congelados, divisiones ni protección:

- **Hoja1**, primera y activa: tabla dinámica nativa `TablaDinámica1`. FACTURACION en A2, filtro inicialmente PROFORMA en B2, encabezados en A4:B4, filas PERTENENCIA2 y valor Cuenta de AFIPOS. Si no existe PROFORMA en el procesamiento, se muestra (Todas), sin categorías ficticias. Si no hay equipos, muestra total cero sin una tabla dinámica vacía inválida.
- **RNC**: AFIPOS, CODIGO_AFILIADO, TERMINAL, NOMBRE_BANCO, MARCA, MODELO, SERIAL, PERTENENCIA, PERTENENCIA2, ESTADO_POSV2, MANEJADOR, FACTURACION, RANGO, en ese orden. Identificadores como texto y autofiltro normal.

La tabla dinámica se alimenta directamente de RNC y contiene los trece campos y registros individuales en una caché nueva, sin datos históricos. Permite filtrar facturación, actualizar y añadir campos como NOMBRE_BANCO. No se exportan tablas auxiliares, fórmulas selectoras ni bloques de validación. Los controles y el análisis detallado se conservan en Streamlit.

RAPIDPAGO/BDV, RAPIDPAGO/BANFANB y SERVICIOS J&G/BDV usan el parque cargado, hoja `7A_PARQUE_PARCIAL_ACTIVOS`: coincidencia exacta de **agente + banco + AFIPOS normalizado** significa COBRO CENTRALIZADO; ausencia significa PROFORMA. No se exige coincidencia adicional de terminal si AFIPOS es válido. Los identificadores numéricos normalizan espacios, sufijo .0 y ceros iniciales para comparar, conservando el texto original en RNC. Nunca se usan coincidencias parciales. Si AFIPOS no sirve, solo se recupera con agente+banco+CODIGO_AFILIADO/AFILIADO+NUMPOS/TERMINAL cuando identifica un único AFIPOS del parque. Si queda vacío, se detiene la exportación para evitar que Cuenta de AFIPOS omita un equipo. El nombre del archivo no bloquea el cruce: el usuario debe cargar el parque correspondiente.

**SITCA se incluye siempre en el alcance comercial y todos sus equipos se clasifican PROFORMA**, incluidas sus equivalencias reconocidas por el diccionario, aunque el agente y banco estén en AGREGADOS EN COBRO o el rango indique transacciones. Las incidencias de calidad de datos siguen visibles en Validaciones. Esta regla del 18 de octubre sustituye su exclusión anterior.

Para POSGLOBAL, una combinación agente+banco incluida en cobro tiene prioridad, transando o sin transacciones. Fuera de cobro, SIN TX es EXONERADO y TRANSANDO es PROFORMA. DISGLOBAL, UBIIPAGOS y LATINPAGOS2021 conservan la exoneración cuando no transan. RANGO determina el estado: TRANSANDO es distinto de SIN TX EN EL MES ACTUAL y 30/60/90/120 DIAS SIN TX. Rangos desconocidos quedan pendientes; las discrepancias con indicadores mensuales se reportan. La regla general sigue siendo agente+banco en cobro → COBRO CENTRALIZADO; fuera → PROFORMA.

Se validan leídos = incluidos + excluidos, resumen = RNC y modalidades = RNC. Las referencias no multiplican filas; se detectan duplicados originales banco+AFIPOS en disco, sin borrarlos silenciosamente.

## Archivos grandes y despliegue

El R34 admite CSV directo o ZIP con **un único CSV**, incluso dentro de una carpeta. Se ignoran otros archivos del ZIP; el diccionario, cobro y parque siguen cargándose por separado. Un ZIP sin CSV, con varios CSV, dañado o protegido con contraseña muestra un error explícito. El CSV comprimido puede tener hasta **4 GB descomprimidos** y sigue sujeto al límite de filas de Excel. Se lee directamente desde el ZIP, sin extraer rutas ni cargar todo el CSV descomprimido en memoria.

El CSV se lee por bloques de 25.000 filas y solo con las columnas necesarias. El Excel se escribe secuencialmente en disco mediante XlsxWriter. La detección de claves repetidas usa SQLite temporal local, sin servicio externo. La previsualización muestra como máximo 1.000 registros.

Streamlit mantiene las subidas y descargas en memoria: un CSV de 626 MB necesita un servidor con margen suficiente. Se recomienda un equipo local o Codespaces con **al menos 4 GB de RAM y varios GB libres de disco temporal**. El límite de subida configurado y versionado en `.streamlit/config.toml` es **1.024 MB por archivo**, tanto CSV como ZIP; la interfaz muestra el límite activo del servidor. En ZIP el límite de subida corresponde al tamaño comprimido; subirlo no incrementa los recursos del servidor. Excel admite como máximo 1.048.575 registros más encabezado en RNC; superar ese límite produce un error explícito.

Codespaces incluye configuración en `.devcontainer`; ejecute `streamlit run app.py` y abra el puerto 8501. Para Streamlit Community Cloud seleccione este repositorio y `app.py`, pero valide su capacidad con el tamaño real. Si el servidor no soporta la carga, ejecute localmente/Codespaces o use la alternativa por archivos en disco:

```sh
python tools/process_files.py --r34 R34.zip --diccionario "DICCIONARIO NUEVO.xlsx" --cobro "AGREGADOS EN COBRO.xlsx" --parque "Parque octubre.xlsx" --salida "RNC Ingresos Actualizado.xlsx"
```

No se registran ni envían archivos de clientes a GitHub. `.gitignore` excluye CSV, XLSX, ZIP y salidas temporales.

## Pruebas y documentación

```sh
python -m unittest discover -s tests -v
```

Incluye reglas comerciales, cruces compuestos, duplicados entre bloques, identificadores largos, filtros, cargas obligatorias, cambios de entrada, conservación de agentes pequeños y ausencia de acumulación entre procesamientos. GitHub Actions ejecuta las pruebas con Python 3.12.

Consulte `docs/analisis-plantilla.md` y `docs/validacion-real.md` para la estructura observada, decisiones y resultados de la prueba completa del ZIP.

La especificación y prueba vigentes están en [docs/correccion-definitiva.md](docs/correccion-definitiva.md). Los informes anteriores documentan requisitos históricos sustituidos por esta corrección.

## Precios e histórico

`config/precios.ejemplo.json` y `agregados/precios.py` preparan la configuración futura de agente, banco, modalidad, precio, moneda, condiciones y vigencia. No hay tarifas inventadas ni cálculo monetario. El procesamiento es una función independiente por archivo; cualquier modo histórico requiere definir claves, actualización y permanencia antes de implementarlo.
