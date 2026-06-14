import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Evaluación de Fórmulas y Nutrientes", layout="wide")

st.title("🧬 Módulo Avanzado: Costos, Precios y Nutrientes")
st.markdown("Carga tu archivo de fórmulas para estructurar materias primas y nutrientes de forma vertical.")

archivo_cargado = st.file_uploader("Selecciona el archivo de Excel de Fórmulas (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        excel_file = pd.ExcelFile(archivo_cargado)
        hojas_disponibles = excel_file.sheet_names
        
        st.sidebar.success("📊 Archivo detectado correctamente")
        
        # --- PROCESAMIENTO DE MATERIAS PRIMAS Y COSTOS ---
        data_formulas = []
        mapeo_hojas = {'PB': 'F.PB', 'Ant': 'F.PYAnterior', 'Act': 'F.PYActual'}
        
        for label, hoja in mapeo_hojas.items():
            if hoja in hojas_disponibles:
                df = pd.read_excel(archivo_cargado, sheet_name=hoja, dtype={'SKU': str, 'Dummy': str})
                id_cols = [c for c in ['Dummy', 'Mes', 'SKU', 'Nombre del producto', 'Proyecto', 'Año'] if c in df.columns]
                mp_cols = [col for col in df.columns if col not in id_cols]
                
                df_v = df.melt(id_vars=id_cols, value_vars=mp_cols, var_name='Cod MP', value_name='Inclusion')
                df_v['Escenario'] = label
                data_formulas.append(df_v)
        
        if data_formulas:
            master_f = pd.concat(data_formulas).dropna(subset=['Inclusion'])
            master_f = master_f[master_f['Inclusion'] > 0]
            
            precio_kpb = pd.DataFrame(columns=['Cod MP', 'Mes', 'Precio'])
            precio_kact = pd.DataFrame(columns=['Cod MP', 'Mes', 'Precio'])
            
            if 'Kardex_PB' in hojas_disponibles:
                precio_kpb = pd.read_excel(archivo_cargado, sheet_name='Kardex_PB', dtype={'Cod MP': str})
            if 'Kardex_Actual' in hojas_disponibles:
                precio_kact = pd.read_excel(archivo_cargado, sheet_name='Kardex_Actual', dtype={'Cod MP': str})
            
            for k_df in [precio_kpb, precio_kact]:
                if 'Código' in k_df.columns: k_df.rename(columns={'Código': 'Cod MP'}, inplace=True)
                if 'Materia prima' in k_df.columns: k_df.rename(columns={'Materia prima': 'Cod MP'}, inplace=True)
            
            df_costos = pd.merge(master_f, precio_kpb[['Cod MP', 'Mes', 'Precio']].drop_duplicates(), on=['Cod MP', 'Mes'], how='left')
            df_costos = pd.merge(df_costos, precio_kact[['Cod MP', 'Mes', 'Precio']].drop_duplicates(), on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))
            
            df_costos['Costo_con_KPB'] = df_costos['Inclusion'] * df_costos.get('Precio_KPB', np.nan)
            df_costos['Costo_con_KACT'] = df_costos['Inclusion'] * df_costos.get('Precio_KACT', np.nan)
            
            df_costos['ALERTA'] = df_costos.apply(lambda x: '🚨 FALTA PRECIO' if pd.isna(x.get('Precio_KACT', np.nan)) else '✅ OK', axis=1)
            
            if 'Act' in df_costos['Escenario'].values and 'PB' in df_costos['Escenario'].values:
                resumen_recetas = df_costos.pivot_table(
                    index=['Dummy', 'Mes'], columns='Escenario', values='Costo_con_KACT', aggfunc='sum'
                ).reset_index()
                if 'Act' in resumen_recetas.columns and 'PB' in resumen_recetas.columns:
                    resumen_recetas['Diff_Act_vs_PB'] = resumen_recetas['Act'] - resumen_recetas['PB']
            else:
                resumen_recetas = pd.DataFrame(columns=['Dummy', 'Mes', 'Mensaje'])

            # --- PROCESAMIENTO VERTICAL DE NUTRIENTES ---
            data_nutrientes = []
            for label, hoja in mapeo_hojas.items():
                hoja_nut = f"Nutrientes_{label}" if f"Nutrientes_{label}" in hojas_disponibles else (hoja if hoja in hojas_disponibles else None)
                if hoja_nut and hoja_nut in hojas_disponibles:
                    df_n = pd.read_excel(archivo_cargado, sheet_name=hoja_nut, dtype={'SKU': str, 'Dummy': str})
                    id_cols_n = [c for c in ['Dummy', 'Mes', 'SKU', 'Nombre del producto', 'Proyecto', 'Año', 'PY'] if c in df_n.columns]
                    nut_cols = [col for col in df_n.columns if col not in id_cols_n]
                    
                    if nut_cols:
                        df_nv = df_n.melt(id_vars=id_cols_n, value_vars=nut_cols, var_name='Cod Nutriente', value_name='Valor Nutriente')
                        df_nv['Escenario'] = label
                        data_nutrientes.append(df_nv)
            
            if data_nutrientes:
                master_nutrientes = pd.concat(data_nutrientes).dropna(subset=['Valor Nutriente'])
                master_nutrientes['Característica'] = master_nutrientes['Cod Nutriente']
            else:
                unique_dummies = master_f['Dummy'].unique() if not master_f.empty else ['DUMMY_REF']
                ejemplo_data = []
                nutrientes_base = {'PROT': 'Proteína Cruda', 'GRASA': 'Grasa Total', 'FIBRA': 'Fibra Cruda', 'HUM': 'Humedad Máxima'}
                for d in unique_dummies:
                    for cod_n, desc_n in nutrientes_base.items():
                        ejemplo_data.append({
                            'Dummy': d, 'Proyecto/PY': 'PY06 JUN', 'Cod Nutriente': cod_n, 'Característica': desc_n, 'Valor Nutriente': 0.0, 'Mes': 'Junio'
                        })
                master_nutrientes = pd.DataFrame(ejemplo_data)

            # --- INTERFAZ ---
            tab1, tab2, tab3 = st.tabs(["🛒 Costos MP", "🧬 Matriz de Nutrientes", "⚠️ Control de Precios"])
            with tab1:
                st.dataframe(df_costos.head(50), use_container_width=True)
            with tab2:
                st.dataframe(master_nutrientes, use_container_width=True)
            with tab3:
                faltantes = df_costos[df_costos['ALERTA'] == '🚨 FALTA PRECIO']
                if not faltantes.empty:
                    st.error("Materias primas sin precio registrado en el Kardex:")
                    st.dataframe(faltantes[['Dummy', 'Mes', 'Cod MP', 'Inclusion']].drop_duplicates(), use_container_width=True)
                else:
                    st.success("✅ Todos los ingredientes tienen precios asignados.")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_costos.to_excel(writer, sheet_name="DATA_PARA_POWERBI", index=False)
                master_nutrientes.to_excel(writer, sheet_name="RELACION_NUTRIENTES", index=False)
                if 'resumen_recetas' in locals() and not resumen_recetas.empty:
                    resumen_recetas.to_excel(writer, sheet_name="RESUMEN_RECETAS", index=False)
            
            procesado_excel = output.getvalue()
            st.sidebar.markdown("---")
            st.sidebar.download_button(
                label="📥 Descargar RESULTADO_COMPARATIVO.xlsx",
                data=procesado_excel,
                file_name="RESULTADO_COMPARATIVO_NUTRIENTES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("💡 Sube tu archivo Excel para procesar.")
