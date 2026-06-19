import pandas as pd
import streamlit as st
import io

st.title("Evaluación de Nutrientes y Recetas - Balanceado")
st.write("Carga tu plantilla mensual para separar de forma independiente Ingredientes (con costos) y Nutrientes.")

# 1. Botón web para cargar el archivo en Streamlit
archivo_cargado = st.file_uploader("Elige tu archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        # Pestañas objetivo
        hojas_targets = {'PB': 'F.PB', 'Ant': 'F.PYAnterior', 'Act': 'F.PYActual'}
        id_cols = ['Dummy', 'Mes', 'SKU', 'Nombre del producto']
        
        data_formulas = []
        data_nutrientes = []
        
        # Leer el archivo cargado
        xl = pd.read_excel(archivo_cargado, sheet_name=None) # Lee todas las hojas de golpe
        
        # 2. Procesamiento Dinámico de Columnas (Melt)
        for label, hoja in hojas_targets.items():
            if hoja in xl.keys():
                df = xl[hoja]
                cols_presentes = [c for c in id_cols if c in df.columns]
                
                # --- SEPARACIÓN DINÁMICA ---
                # Ingredientes: Todo lo que empiece con 'PC_'
                cols_ingredientes = [c for c in df.columns if str(c).startswith('PC_')]
                # Nutrientes: Todo lo que NO sea ID ni sea ingrediente
                cols_nutrientes = [c for c in df.columns if c not in cols_presentes and c not in cols_ingredientes]
                
                # Trasponer Ingredientes
                if cols_ingredientes:
                    df_ing_v = df.melt(id_vars=cols_presentes, value_vars=cols_ingredientes, var_name='Cod MP', value_name='Inclusion')
                    df_ing_v['Escenario'] = label
                    data_formulas.append(df_ing_v)
                    
                # Trasponer Nutrientes
                if cols_nutrientes:
                    df_nut_v = df.melt(id_vars=cols_presentes, value_vars=cols_nutrientes, var_name='Cod Nutriente', value_name='Valor Nutriente')
                    df_nut_v['Escenario'] = label
                    data_nutrientes.append(df_nut_v)

        # ==========================================
        # SECCIÓN INGREDIENTES + PRECIOS (Tu base original)
        # ==========================================
        df_costos = pd.DataFrame()
        resumen_recetas = pd.DataFrame()
        
        if data_formulas:
            master_f = pd.concat(data_formulas).dropna(subset=['Inclusion'])
            
            if 'Kardex_PB' in xl.keys() and 'Kardex_Actual' in xl.keys():
                k_pb = xl['Kardex_PB']
                k_act = xl['Kardex_Actual']
                
                # Forzar formato texto para códigos
                master_f['Cod MP'] = master_f['Cod MP'].astype(str).str.strip()
                k_pb['Cod MP'] = k_pb['Cod MP'].astype(str).str.strip()
                k_act['Cod MP'] = k_act['Cod MP'].astype(str).str.strip()
                
                # Cruce de precios
                df_costos = pd.merge(master_f, k_pb[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left')
                df_costos = pd.merge(df_costos, k_act[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))
                
                # Cálculos
                df_costos['Costo_con_KPB'] = df_costos['Inclusion'] * df_costos['Precio_KPB']
                df_costos['Costo_con_KACT'] = df_costos['Inclusion'] * df_costos['Precio_KACT']
                df_costos['ALERTA'] = df_costos.apply(lambda x: 'FALTA PRECIO' if pd.isna(x['Precio_KACT']) else 'OK', axis=1)
                
                # Pivot resumen de recetas
                resumen_recetas = df_costos.pivot_table(index=['Dummy', 'Mes'], columns='Escenario', values='Costo_con_KACT', aggfunc='sum')
                if 'Act' in resumen_recetas.columns and 'PB' in resumen_recetas.columns:
                    resumen_recetas['Diff_Act_vs_PB'] = resumen_recetas['Act'] - resumen_recetas['PB']
            else:
                df_costos = master_f

        # ==========================================
        # SECCIÓN NUTRIENTES
        # ==========================================
        master_nutrientes = pd.DataFrame()
        if data_nutrientes:
            master_nutrientes = pd.concat(data_nutrientes).dropna(subset=['Valor Nutriente'])

        # ==========================================
        # EXPORTACIÓN Y DESCARGA (Formato Web Streamlit)
        # ==========================================
        # Validar que al menos tengamos un dato para no generar el error de pestañas visibles
        if not df_costos.empty or not master_nutrientes.empty:
            st.success("¡Datos procesados correctamente!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_costos.empty:
                    df_costos.to_excel(writer, sheet_name="DATA_PARA_POWERBI", index=False)
                if not resumen_recetas.empty:
                    resumen_recetas.to_excel(writer, sheet_name="RESUMEN_RECETAS")
                if not df_costos.empty and 'ALERTA' in df_costos.columns:
                    df_costos[df_costos['ALERTA'] == 'FALTA PRECIO'].to_excel(writer, sheet_name="ALARMAS", index=False)
                if not master_nutrientes.empty:
                    master_nutrientes.to_excel(writer, sheet_name="VERTICAL_NUTRIENTES", index=False)
            
            output.seek(0)
            
            # Botón de descarga web
            st.download_button(
                label="📥 Descargar Reporte Consolidado",
                data=output,
                file_name="RESULTADO_COMPARATIVO_CON_NUTRIENTES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("El archivo cargado no contiene las pestañas requeridas (F.PB, F.PYAnterior, F.PYActual).")

    except Exception as e:
        st.error(f"Error crítico en el análisis: {e}")
        
