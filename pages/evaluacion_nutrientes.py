import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Evaluación de Fórmulas y Nutrientes", layout="wide")

st.title("🧪 Módulo de Separación Avanzada: Ingredientes y Nutrientes")
st.markdown("Carga tu archivo de formulación. El sistema separará la información en dos hojas independientes manteniendo el alineamiento exacto por columnas (Dummy / Proyecto).")

archivo_cargado = st.file_uploader("Selecciona el archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        excel_file = pd.ExcelFile(archivo_cargado)
        hojas_disponibles = excel_file.sheet_names
        
        # Filtrar pestañas operativas (omitir hojas de Kardex si las hay para buscar solo las matrices de fórmulas)
        pestanas_formulas = [h for h in hojas_disponibles if 'kardex' not in h.lower()]
        
        st.sidebar.success(f"📊 Pestañas encontradas: {pestanas_formulas}")
        
        data_formulas_TOTAL = []
        data_nutrientes_TOTAL = []
        
        for hoja in pestanas_formulas:
            # Lectura inicial cruda para escanear las coordenadas de las tablas
            df_crudo = pd.read_excel(archivo_cargado, sheet_name=hoja, header=None)
            
            fila_inicio_mp = None
            fila_inicio_nut = None
            
            # Localizar dinámicamente dónde empieza el bloque de Composición y dónde el de Análisis
            for idx, row in df_crudo.iterrows():
                row_str = [str(x).strip() for x in row.values]
                if "Composición" in row_str or ("Cód." in row_str and "Materia prima" in row_str):
                    if fila_inicio_mp is None:
                        fila_inicio_mp = idx
                if "Análisis" in row_str or "Tipo de característica" in row_str:
                    fila_inicio_nut = idx
                    break
            
            # --- PROCESAMIENTO HOJA 1: MATERIAS PRIMAS + COSTO ---
            if fila_inicio_mp is not None:
                # Extraer metadatos de la parte superior (Año, Mes, Grupo, Carpeta)
                # Buscamos valores de referencia en las primeras filas
                grupo_val = df_crudo.iloc[1, 1] if df_crudo.shape[1] > 1 else hoja
                establecimiento = df_crudo.iloc[2, 1] if df_crudo.shape[1] > 1 else ""
                
                df_mp = pd.read_excel(archivo_cargado, sheet_name=hoja, skiprows=fila_inicio_mp + 1)
                
                # Recortar la tabla justo antes de que empiece la sección de Nutrientes
                if fila_inicio_nut is not None:
                    df_mp = df_mp.iloc[:(fila_inicio_nut - fila_inicio_mp - 3)]
                
                # Limpieza de filas de totales o vacías
                df_mp = df_mp.dropna(subset=['Cód.', 'Materia prima'], how='all')
                df_mp = df_mp[df_mp['Cód.'] != 'Total']
                
                id_cols_mp = ['Cód.', 'Materia prima', 'Precio']
                columnas_productos = [c for c in df_mp.columns if c not in id_cols_mp and 'Unnamed' not in str(c)]
                
                # Desdinamizar (Melt) para volverlo estructurado
                df_mp_v = df_mp.melt(id_vars=id_cols_mp, value_vars=columnas_productos, var_name='Dummy', value_name='Peso / Inclusión')
                
                # Inyectar metadatos del proyecto y mes
                df_mp_v['Proyecto'] = grupo_val
                df_mp_v['Establecimiento'] = establecimiento
                df_mp_v['Mes'] = "Junio"  # Puede automatizarse detectando la celda específica de tu plantilla
                
                # NUEVO: Cálculo e inclusión del costo solicitado
                df_mp_v['Costo_Calculado'] = df_mp_v['Peso / Inclusión'] * df_mp_v['Precio']
                
                # Reordenar columnas para mantener tu esquema clásico
                columnas_mp_final = ['Dummy', 'Proyecto', 'Establecimiento', 'Mes', 'Cód.', 'Materia prima', 'Precio', 'Peso / Inclusión', 'Costo_Calculado']
                data_formulas_TOTAL.append(df_mp_v[columnas_mp_final])
                
            # --- PROCESAMIENTO HOJA 2: NUTRIENTES CON ENCABEZADO ESPEJO ---
            if fila_inicio_nut is not None:
                df_nut = pd.read_excel(archivo_cargado, sheet_name=hoja, skiprows=fila_inicio_nut + 1)
                df_nut = df_nut.dropna(subset=['Cód.', 'Caracterís'], how='all')
                
                id_cols_nut = ['Tipo de característica', 'Cód.', 'Caracterís', 'Unidad']
                id_cols_existentes = [c for c in id_cols_nut if c in df_nut.columns]
                
                columnas_valores_nut = [c for c in df_nut.columns if c not in id_cols_existentes and 'Unnamed' not in str(c)]
                
                # Melt de Nutrientes
                df_nut_v = df_nut.melt(id_vars=id_cols_existentes, value_vars=columnas_valores_nut, var_name='Columna_Index', value_name='Valor Nutriente')
                
                # MAREAR ALINEACIÓN EN ESPEJO: Asociar el mismo Dummy/Fórmula según la posición de la columna
                if len(columnas_productos) > 0:
                    mapeo_espejo = dict(zip(columnas_valores_nut, columnas_productos))
                    df_nut_v['Dummy'] = df_nut_v['Columna_Index'].map(mapeo_espejo)
                else:
                    df_nut_v['Dummy'] = df_nut_v['Columna_Index']
                
                df_nut_v['Proyecto'] = grupo_val
                df_nut_v['Mes'] = "Junio"
                
                # Renombrar columnas para claridad del reporte
                df_nut_v = df_nut_v.rename(columns={'Cód.': 'Cod Nutriente', 'Caracterís': 'Característica'})
                
                columnas_nut_final = ['Dummy', 'Proyecto', 'Mes', 'Tipo de característica', 'Cod Nutriente', 'Característica', 'Unidad', 'Valor Nutriente']
                columnas_nut_existentes = [c for c in columnas_nut_final if c in df_nut_v.columns]
                data_nutrientes_TOTAL.append(df_nut_v[columnas_nut_existentes])

        # --- CONSOLIDACIÓN DE TABLAS ---
        if data_formulas_TOTAL:
            df_hoja1_ingredientes = pd.concat(data_formulas_TOTAL).dropna(subset=['Peso / Inclusión'])
            df_hoja1_ingredientes = df_hoja1_ingredientes[df_hoja1_ingredientes['Peso / Inclusión'] > 0]
        else:
            df_hoja1_ingredientes = pd.DataFrame()
            
        if data_nutrientes_TOTAL:
            df_hoja2_nutrientes = pd.concat(data_nutrientes_TOTAL).dropna(subset=['Valor Nutriente'])
        else:
            df_hoja2_nutrientes = pd.DataFrame()

        # --- VISTA PREVIA INTERACTIVA EN STREAMLIT ---
        tab1, tab2 = st.tabs(["🛒 Hoja 1: Ingredientes + Costos", "🧬 Hoja 2: Nutrientes (Espejo)"])
        
        with tab1:
            st.subheader("Estructura de Fórmulas con Inclusión de Costo")
            st.dataframe(df_hoja1_ingredientes, use_container_width=True)
            
        with tab2:
            st.subheader("Estructura de Análisis Nutricional con Encabezado de Producto (Dummy)")
            st.dataframe(df_hoja2_nutrientes, use_container_width=True)

        # --- GENERAR EXCEL FINAL CON LAS DOS PESTAÑAS SEPARADAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not df_hoja1_ingredientes.empty:
                df_hoja1_ingredientes.to_excel(writer, sheet_name="RESUMEN_INGREDIENTES_Y_COSTOS", index=False)
            if not df_hoja2_nutrientes.empty:
                df_hoja2_nutrientes.to_excel(writer, sheet_name="ANALISIS_NUTRIENTES_ESPEJO", index=False)
        
        excel_final = output.getvalue()
        
        st.sidebar.markdown("---")
        st.sidebar.download_button(
            label="📥 Descargar Reporte Multihoja",
            data=excel_final,
            file_name="EVALUACION_PROYECTOS_COMPLETA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error procesando las coordenadas del archivo: {e}")
else:
    st.info("💡 Carga el archivo comprimido o individual de fórmulas para procesar ambas hojas en espejo.")
