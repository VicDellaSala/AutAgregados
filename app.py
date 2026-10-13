from pathlib import Path
from tempfile import TemporaryDirectory
import streamlit as st
from agregados.procesamiento import process
from agregados.normalizacion import MODES

st.set_page_config(page_title='Automatización de Agregados Comerciales', page_icon='📊', layout='wide')
st.title('Automatización de Agregados Comerciales')
st.write('Cargue los cuatro archivos para generar el RNC y su resumen. Se procesan todos los registros del R34, sin seleccionar un mes.')


def clear_result():
    for key in list(st.session_state):
        if key.startswith('filter_'):
            del st.session_state[key]
    old = st.session_state.pop('result', None)
    if old:
        old['temp'].cleanup()


with st.expander('Opciones de lectura del archivo'):
    encoding = st.selectbox('Codificación del CSV', ['utf-8-sig', 'cp1252', 'latin1'], on_change=clear_result)
    st.caption(f"Límite de subida: {st.get_option('server.maxUploadSize'):,} MB por archivo. La capacidad de procesamiento depende de la memoria del servidor.")

left, right = st.columns(2)
with left:
    r34 = st.file_uploader('1. R34 (CSV o ZIP)', type=['csv', 'zip'], on_change=clear_result, help='El ZIP debe contener un único CSV del R34, incluso dentro de una carpeta. Los demás archivos se ignoran. Límite del CSV dentro del ZIP: 4 GB. Se procesa por bloques sin extraerlo completo en memoria.')
    dictionary = st.file_uploader('2. DICCIONARIO NUEVO', type=['xlsx'], on_change=clear_result)
with right:
    billing = st.file_uploader('3. AGREGADOS EN COBRO', type=['xlsx'], on_change=clear_result)
    park = st.file_uploader('4. PARQUE PARCIAL', type=['xlsx'], on_change=clear_result)
st.caption('RNC ingreso.xlsx es solo el modelo de diseño: no debe cargarse. Las excepciones del parque se comprueban con las fechas del R34 cuando están disponibles.')

if st.button('Procesar archivos', type='primary', disabled=any(f is None for f in [r34, dictionary, billing, park])):
    clear_result()
    temp = TemporaryDirectory(prefix='rnc-streamlit-')
    status = st.empty()
    try:
        with st.spinner('Procesando y generando el Excel…'):
            result = process(r34, dictionary, billing, park, Path(temp.name)/'RNC Ingresos Actualizado.xlsx', encoding=encoding,
                             progress=lambda n: status.info(f'{n:,} registros leídos. Generando RNC…'))
        st.session_state.result = {'data': result, 'temp': temp}
        status.success('Archivo generado. El resumen y el detalle cuadran.')
    except Exception as error:
        temp.cleanup()
        status.empty()
        st.error(f'No se generó el archivo: {error}')

if 'result' in st.session_state:
    result = st.session_state.result['data']
    audit = result.audit
    cols = st.columns(4)
    for column, label, number in zip(cols, ['Equipos en RNC', 'Excluidos', 'Pendientes', 'Exonerados'], [audit.processed, sum(audit.excluded.values()), audit.modes.get(MODES[3], 0), audit.modes.get(MODES[2], 0)]):
        column.metric(label, f'{number:,}' if isinstance(number, int) else number)
    if audit.modes.get(MODES[3], 0):
        st.warning('Hay registros pendientes de validación. Permanecen identificados en RNC y en el resumen, separados de las modalidades resueltas.')
    tab_rnc, tab_summary, tab_validation = st.tabs(['RNC', 'RESUMEN', 'Validaciones'])
    with tab_rnc:
        st.caption('Vista previa de hasta 1.000 registros. El Excel contiene todos los registros aplicables.')
        st.dataframe(result.preview, hide_index=True, width='stretch')
    with tab_summary:
        data = result.summary
        for column, label in [('FACTURACION', 'Facturación'), ('RANGO', 'Rango'), ('AGENTE', 'Agente agregado'), ('BANCO', 'Banco')]:
            choices = st.multiselect(label, sorted(data[column].unique()), key='filter_'+column)
            if choices:
                data = data.loc[data[column].isin(choices)]
        st.metric('Total general filtrado', f'{int(data.EQUIPOS.sum()):,}')
        st.caption('Agentes con sus bancos. Se incluyen subtotales por agente y cantidades por modalidad.')
        if len(data):
            grouped = data.pivot_table(index=['AGENTE', 'BANCO'], columns='FACTURACION', values='EQUIPOS', aggfunc='sum', fill_value=0).reindex(columns=MODES, fill_value=0)
            grouped['Cuenta de AFIPOS'] = grouped.sum(axis=1)
            st.dataframe(grouped, width='stretch')
            st.write('Subtotales por agente')
            st.dataframe(grouped.groupby(level='AGENTE').sum(), width='stretch')
        st.caption('En Excel se conserva el resumen con fórmulas y se añade una tabla dinámica nativa al final de RESUMEN. Use el enlace «Tabla dinámica nativa». Su base es el detalle agrupado H:L; el doble clic muestra grupos, no equipos individuales.')
    with tab_validation:
        st.dataframe([{'Control': k, 'Valor': v} for k, v in audit.controls(int(result.summary.EQUIPOS.sum())).items()], hide_index=True)
        for note in audit.notes:
            st.info(note)
        for title, values in [('Exclusiones por agente', audit.excluded), ('Pertenencias no reconocidas', audit.unknown), ('Incidencias', audit.issues)]:
            st.write(title)
            st.dataframe([{'Motivo / pertenencia': k, 'Registros': v} for k, v in values.items()], hide_index=True)
        st.write('Muestra de hasta 300 incidencias')
        st.dataframe(audit.samples, hide_index=True)
    with result.output.open('rb') as download:
        st.download_button('Descargar RNC Ingresos Actualizado.xlsx', download, file_name='RNC Ingresos Actualizado.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
