import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Evaluación de Fórmulas y Nutrientes", layout="wide")

st.title("🧪 Módulo de Separación Avanzada: Ingredientes y Nutrientes")
st.markdown("Carga tu archivo de formulación. El sistema detectará las secciones de Composición y Análisis de forma dinámica.")

archivo_cargado = st.file_uploader("Selecciona el archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        excel_file = pd.ExcelFile(archivo_cargado)
        hojas_disponibles = excel_file.sheet_names
        
        # Omitir hojas de Kardex o configuraciones
        pestanas_formulas = [h for h in hojas_disponibles if 'kardex' not in h.lower() and 'config' not in h.lower()]
        
        st.sidebar.success(f"📊 Pestañas detectadas: {pestanas_formulas}")
        
        data_formulas_TOTAL = []
        data_nutrientes_TOTAL = []
        
        for hoja in pestanas_formulas:
            df_crudo = pd.read_excel(archivo_cargado, sheet_name=hoja, header=None)
            
            fila_inicio_mp = None
            fila_inicio_nut = None
            
            # Buscar dinámicamente las palabras clave de inicio
            for idx, row in df_crudo.iterrows():
                row_str = [str(x).strip().lower() for x in row.values]
                if "composición" in row_str or "composicion" in row_str or ("cód." in row_str and "materia prima" in row_str):
                    if fila_inicio_mp is None:
                        fila_inicio_mp = idx
                if "análisis" in row_str or "analisis" in row_str or "tipo de característica" in row_str or "tipo de caracteristica" in row_str:
                    fila_inicio_nut = idx
                    break
            
            if fila_inicio_mp is None:
                continue
                
            # --- DETECTAR ENCABEZADOS DE PRODUCTO (DUMMY) ---
            # Escaneamos las filas previas a "Composición" para hallar la fila que tiene los códigos de producto o fórmulas
            fila_cabecera_productos = fila_inicio_mp - 1
            for r in range(max(0, fila_inicio_mp - 8), fila_inicio_mp):
                valores_fila = [str(x).strip() for x in df_crudo.iloc[r].values if pd.notna(x)]
                # Si encontramos las celdas con códigos típicos de la planta (ej: HD14496 o números de optimización)
                if any(x.startswith('HD') or x.isdigit() for x in valores_fila):
                    fila_cabecera_productos = r
                    break

            # Extraer metadatos generales
            grupo_val = df_crudo.iloc[1, 1] if df_crudo.shape[1] > 1 else hoja
            establecimiento = df_crudo.iloc[2, 1] if df_crudo.shape[1] > 1 else "INB"
            mes_val = df_crudo.iloc[1, 3] if df_crudo.shape[1] > 3 else "Junio"
            
            # --- PROCESAR SECCIÓN 1: MATERIAS PRIMAS ---
            df_mp = pd.read_excel(archivo_cargado, sheet_name=hoja, skiprows=fila_inicio_mp)
            if fila_inicio_nut is not None:
                # Cortar la tabla antes de que inicien los nutrientes
                limite = fila_inicio_nut - fila_inicio_mp - 1
                df_mp = df_mp.iloc[:limite]
            
            # Reasignar la fila correcta de productos si se detectó una fila superior
            if fila_cabecera_productos != fila_inicio_mp:
                cabecera_real = pd.read_excel(archivo_cargado, sheet_name=hoja, skiprows=fila_cabecera_productos)
                if fila_inicio_nut is not None:
                    cabecera_real = cabecera_real.iloc[:(fila_inicio_nut - fila_cabecera_productos - 1)]
                
                # Sincronizar nombres de columnas manteniendo Cód., Materia prima y Precio
                for col in cabecera_real.columns:
                    if col not in ['Cód.', 'Materia prima', 'Precio'] and not str(col).startswith('Unnamed'):
                        # Buscar la misma columna en df_mp y renombrarla
                        idx_col = cabecera_real.columns.get_loc(col)
                        if idx_col < len(df_mp.columns):
                            df_mp.rename(columns={df_mp.columns[idx_col]: col}, inplace=True)

            df_mp = df_mp.dropna(subset=['Cód.', 'Materia prima'], how='all')
            df_mp = df_mp[~df_mp['Cód.'].astype(str).str.contains('Total|Composición|composicion', case=False, na=False)]
            
            id_cols_mp = [c for c in ['Cód.', 'Materia prima', 'Precio'] if c in df_mp.columns]
            columnas_productos = [c for c in df_mp.columns if c not in id_cols_mp and not str(c).startswith('Unnamed')]
            
            # Melt de Ingredientes
            df_mp_v = df_mp.melt(id_vars=id_cols_mp, value_vars=columnas_productos, var_name='Dummy', value_name='Peso / Inclusión')
            df_mp_v['Proyecto'] = grupo_val
            df_mp_v['Establecimiento'] = establecimiento
            df_mp_v['Mes'] = mes_val
            
            # Calcular Costo
            df_mp_v['Precio'] = pd.to_numeric(df_mp_v['Precio'], errors='coerce').fillna(0)
            df_mp_v['Peso / Inclusión'] = pd.to_numeric(df_mp_v['Peso / Inclusión'], errors='coerce').fillna(0)
            df_mp_v['Costo_Calculado'] = df_mp_v['Peso / Inclusión'] * df_mp_v['Precio']
            
            columnas_mp_final = ['Dummy', 'Proyecto', 'Establecimiento', 'Mes', 'Cód.', 'Materia prima', 'Precio', 'Peso / Inclusión', 'Costo_Calculado']
            columnas_mp_existentes = [c for c in columnas_mp_final if c in df_mp_v.columns]
            data_formulas_TOTAL.append(df_mp_v[columnas_mp_existentes])
            
            # --- PROCESAR SECCIÓN 2: NUTRIENTES (ESPEJO) ---
            if fila_inicio_nut is not None:
                df_nut = pd.read_excel(archivo_cargado, sheet_name=hoja, skiprows=fila_inicio_nut)
                df_nut = df_nut.dropna(subset=['Cód.', 'Caracterís'], how='all', errors='ignore')
                if 'Cód.' not in df_nut.columns and 'Cod.' in df_nut.columns:
                    df_nut.rename(columns={'Cod.': 'Cód.'}, inplace=True)
                if 'Caracterís' not in df_nut.columns and 'Característica' in df_nut.columns:
                    df_nut.rename(columns={'Característica': 'Caracterís'}, inplace=True)
                
                id_cols_nut = ['Tipo de característica', 'Cód.', 'Caracterís', 'Unidad']
                id_cols_existentes = [c for c in id_cols_nut if c in df_nut.columns]
                
                columnas_valores_nut = [c for c in df_nut.columns if c not in id_cols_existentes and not str(c).startswith('Unnamed')]
                
                df_nut_v = df_nut.melt(id_vars=id_cols_existentes, value_vars=columnas_valores_nut, var_name='Columna_Index', value_name='Valor Nutriente')
                
                # Alinear en espejo con las columnas de productos identificadas arriba
                if len(columnas_productos) > 0:
                    mapeo_espejo = dict(zip(columnas_valores_nut, columnas_productos))
                    df_nut_v['Dummy'] = df_nut_v['Columna_Index'].map(mapeo_espejo)
                else:
                    df_nut_v['Dummy'] = df_nut_v['Columna_Index']
                
                df_nut_v['Proyecto'] = grupo_val
                df_nut_v['Mes'] = mes_val
                
                df_nut_v = df_nut_v.rename(columns={'Cód.': 'Cod Nutriente', 'Caracterís': 'Característica'})
                
                columnas_nut_final = ['Dummy', 'Proyecto', 'Mes', 'Tipo de característica', 'Cod Nutriente', 'Característica', 'Unidad', 'Valor Nutriente']
                columnas_nut_existentes = [c for c in columnas_nut_final if c in df_nut_v.columns]
                data_nutrientes_TOTAL.append(df_nut_v[columnas_nut_existentes])

        # --- CONSOLIDACIÓN ---
        if data_formulas_TOTAL:
            df_hoja1 = pd.concat(data_formulas_TOTAL).dropna(subset=['Peso / Inclusión'])
            df_hoja1 = df_hoja1[df_hoja1['Peso / Inclusión'] > 0]
        else:
            df_hoja1 = pd.DataFrame()
            
        if data_nutrientes_TOTAL:
            df_hoja2 = pd.concat(data_nutrientes_TOTAL).dropna(subset=['Valor Nutriente'])
        else:
            df_hoja2 = pd.DataFrame()

        # --- MOSTRAR EN PANTALLA ---
        tab1, tab2 = st.tabs(["🛒 Hoja 1: Ingredientes + Costos", "🧬 Hoja 2: Nutrientes (Espejo)"])
        with tab1:
            st.dataframe(df_hoja1, use_container_width=True)
        with tab2:
            st.dataframe(df_hoja2, use_container_width=True)

        # --- EXCEL MULTIHOJA ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not df_hoja1.empty:
                df_hoja1.to_excel(writer, sheet_name="RESUMEN_INGREDIENTES_Y_COSTOS", index=False)
            if not df_hoja2.empty:
                df_hoja2.to_excel(writer, sheet_name="ANALISIS_NUTRIENTES_ESPEJO", index=False)
        
        excel_final = output.getvalue()
        st.sidebar.markdown("---")
        st.sidebar.download_button(
            label="📥 Descargar Reporte Multihoja",
            data=excel_final,
            file_name="EVALUACION_PROYECTOS_COMPLETA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
else:
    st.info("👋 Sube tu archivo de fórmulas para separar ingredientes y nutrientes de forma automática.")
