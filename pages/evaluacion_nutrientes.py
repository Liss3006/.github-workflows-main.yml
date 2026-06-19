import pandas as pd
import streamlit as st
import io

st.title("Transformador Masivo: Ingredientes y Nutrientes")
st.write("Carga cualquier plantilla horizontal para separar automáticamente los ingredientes y el perfil nutricional en pestañas independientes.")

# 1. Carga del archivo de Excel
archivo_cargado = st.file_uploader("Elige tu archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        id_cols_base = ['Dummy', 'Mes', 'SKU', 'Nombre del producto']
        
        data_ingredientes_total = []
        data_nutrientes_total = []
        
        # Leer todas las pestañas del archivo de forma dinámica
        xl = pd.read_excel(archivo_cargado, sheet_name=None)
        
        # Guardamos las pestañas de Kardex si existen para el cruce de precios posterior
        k_pb = xl.get('Kardex_PB', None)
        k_act = xl.get('Kardex_Actual', None)
        
        for nombre_hoja, df in xl.items():
            # Ignorar las pestañas de Kardex o de resultados previos durante el mapeo masivo
            if nombre_hoja in ['Kardex_PB', 'Kardex_Actual', 'DATA_PARA_POWERBI', 'RESUMEN_RECETAS', 'ALARMAS', 'INGREDIENTES', 'NUTRIENTES']:
                continue
                
            # Validar qué columnas de identificación están presentes en esta pestaña
            cols_presentes = [c for c in id_cols_base if c in df.columns]
            if not cols_presentes:
                # Si no encuentra las columnas estándar, toma las primeras 4 como identificación por seguridad
                cols_presentes = list(df.columns[:4])
            
            # --- SEPARACIÓN DINÁMICA DE COLUMNAS ---
            # Ingredientes: Columnas que empiezan con 'PC_'
            cols_ingredientes = [c for c in df.columns if str(c).startswith('PC_')]
            # Nutrientes: Columnas que NO son de identificación ni empiezan con 'PC_'
            cols_nutrientes = [c for c in df.columns if c not in cols_presentes and c not in cols_ingredientes]
            
            # Trasponer bloque de Ingredientes (Naranja)
            if cols_ingredientes:
                df_ing_v = df.melt(
                    id_vars=cols_presentes, 
                    value_vars=cols_ingredientes, 
                    var_name='Cod MP', 
                    value_name='Inclusion'
                )
                df_ing_v['Origen_Pestaña'] = nombre_hoja
                data_ingredientes_total.append(df_ing_v)
                
            # Trasponer bloque de Nutrientes (Gris)
            if cols_nutrientes:
                df_nut_v = df.melt(
                    id_vars=cols_presentes, 
                    value_vars=cols_nutrientes, 
                    var_name='Cod Nutriente', 
                    value_name='Valor Nutriente'
                )
                df_nut_v['Origen_Pestaña'] = nombre_hoja
                data_nutrientes_total.append(df_nut_v)

        # ==========================================
        # CONSTRUCCIÓN DE LA PESTAÑA INGREDIENTES
        # ==========================================
        df_ingredientes_final = pd.DataFrame()
        if data_ingredientes_total:
            master_ing = pd.concat(data_ingredientes_total).dropna(subset=['Inclusion'])
            
            # Si el archivo incluye las pestañas de Kardex, hacemos el cruce de precios automáticamente
            if k_pb is not None and k_act is not None:
                master_ing['Cod MP'] = master_ing['Cod MP'].astype(str).str.strip()
                k_pb['Cod MP'] = k_pb['Cod MP'].astype(str).str.strip()
                k_act['Cod MP'] = k_act['Cod MP'].astype(str).str.strip()
                
                # Mapeo de precios por código y mes
                df_ingredientes_final = pd.merge(master_ing, k_pb[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left')
                df_ingredientes_final = pd.merge(df_ingredientes_final, k_act[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))
                
                # Cálculos económicos adicionales
                df_ingredientes_final['Costo_con_KPB'] = df_ingredientes_final['Inclusion'] * df_ingredientes_final['Precio_KPB']
                df_ingredientes_final['Costo_con_KACT'] = df_ingredientes_final['Inclusion'] * df_ingredientes_final['Precio_KACT']
                df_ingredientes_final['ALERTA'] = df_ingredientes_final.apply(lambda x: 'FALTA PRECIO' if pd.isna(x['Precio_KACT']) else 'OK', axis=1)
            else:
                df_ingredientes_final = master_ing

        # ==========================================
        # CONSTRUCCIÓN DE LA PESTAÑA NUTRIENTES
        # ==========================================
        df_nutrientes_final = pd.DataFrame()
        if data_nutrientes_total:
            df_nutrientes_final = pd.concat(data_nutrientes_total).dropna(subset=['Valor Nutriente'])

        # ==========================================
        # GENERACIÓN DEL ARCHIVO CON LAS DOS PESTAÑAS
        # ==========================================
        if not df_ingredientes_final.empty or not df_nutrientes_final.empty:
            st.success("¡Estructura procesada y separada exitosamente!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Pestaña 1: Solo Ingredientes (con columnas fijas, códigos, inclusiones y precios si aplica)
                if not df_ingredientes_final.empty:
                    df_ingredientes_final.to_excel(writer, sheet_name="INGREDIENTES", index=False)
                    st.write("### Vista previa de pestaña INGREDIENTES:")
                    st.dataframe(df_ingredientes_final.head(5))
                
                # Pestaña 2: Solo Nutrientes (con columnas fijas, códigos de nutrientes y sus valores)
                if not df_nutrientes_final.empty:
                    df_nutrientes_final.to_excel(writer, sheet_name="NUTRIENTES", index=False)
                    st.write("### Vista previa de pestaña NUTRIENTES:")
                    st.dataframe(df_nutrientes_final.head(5))
            
            output.seek(0)
            
            # Botón único de descarga
            st.download_button(
                label="📥 Descargar Formato Consolidado",
                data=output,
                file_name="CONSOLIDADO_INGREDIENTES_Y_NUTRIENTES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron datos estructurados válidos para procesar en el archivo.")

    except Exception as e:
        st.error(f"Error durante el procesamiento del archivo: {e}")
