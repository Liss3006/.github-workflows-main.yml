import pandas as pd
import streamlit as st
import io

st.title("Módulo de Separación Avanzada: Sincronización de Cabeceras")
st.markdown("Carga tu archivo de formulación. Las cabeceras e índices se alinearán simétricamente para Ingredientes y Nutrientes.")

def expandir_meses(rango_texto):
    meses_ano = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    meses_map = {m[:3].lower(): m for m in meses_ano}
    
    texto = str(rango_texto).strip()
    texto_limpio = texto.replace('PB', '').replace('Actual', '').strip()
    
    if '-' in texto_limpio:
        partes = texto_limpio.split('-')
        inicio = partes[0].strip()[:3].lower()
        fin = partes[1].strip()[:3].lower()
        
        if inicio in meses_map and fin in meses_map:
            idx_inicio = meses_ano.index(meses_map[inicio])
            idx_fin = meses_ano.index(meses_map[fin])
            if idx_inicio <= idx_fin:
                return meses_ano[idx_inicio:idx_fin+1]
            else:
                return meses_ano[idx_inicio:] + meses_ano[:idx_fin+1]
                
    for corto, largo in meses_map.items():
        if corto in texto_limpio.lower():
            return [largo]
            
    return [texto]

archivo_cargado = st.file_uploader("Selecciona el archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        xl = pd.ExcelFile(archivo_cargado)
        
        data_ingredientes_total = []
        data_nutrientes_total = []
        
        # Mapear hojas de Kardex si existen para traer precios
        k_pb = pd.read_excel(archivo_cargado, sheet_name='Kardex_PB') if 'Kardex_PB' in xl.sheet_names else None
        k_act = pd.read_excel(archivo_cargado, sheet_name='Kardex_Actual') if 'Kardex_Actual' in xl.sheet_names else None
        
        for hoja in xl.sheet_names:
            if hoja in ['Kardex_PB', 'Kardex_Actual', 'DATA_PARA_POWERBI', 'RESUMEN_RECETAS', 'ALARMAS', 'INGREDIENTES', 'NUTRIENTES']:
                continue
                
            df = pd.read_excel(archivo_cargado, sheet_name=hoja, header=None)
            
            # --- DETECTAR SECCIONES DINÁMICAMENTE ---
            idx_composicion = None
            idx_analisis = None
            idx_total_kg = None
            
            for idx, row in df.iterrows():
                val_a = str(row[0]).strip().upper()
                val_b = str(row[1]).strip().upper()
                if 'COMPOSICIÓN' in val_a or 'COMPOSICIÓN' in val_b or 'COMPOSICION' in val_a or 'COMPOSICION' in val_b:
                    idx_composicion = idx
                if 'ANÁLISIS' in val_a or 'ANÁLISIS' in val_b or 'ANALISIS' in val_a or 'ANALISIS' in val_b:
                    idx_analisis = idx
                if 'TOTAL, KG' in val_b or 'TOTAL' in val_b:
                    idx_total_kg = idx

            if idx_composicion is None or idx_analisis is None:
                continue
                
            fin_ingredientes = idx_total_kg if idx_total_kg is not None else idx_analisis
            
            # --- CONSTRUIR DICCIONARIO MAESTRO DE CABECERAS POR COLUMNA REAL ---
            # Guardaremos la metadata indexada por la posición de la columna para evitar desfases
            dict_cabeceras = {}
            
            # Fila 5 (Estabs/Escenarios), Fila 6 (Cód. Dummies), Fila 8 (Carpetas/Meses)
            for col in range(2, df.shape[1]):
                raw_dummy = str(df.iloc[6, col]).strip()
                
                # Si no hay un código válido en la fila 7, no es una columna de datos real
                if raw_dummy == 'nan' or not raw_dummy or 'CÓD' in raw_dummy.upper():
                    continue
                
                raw_esc = str(df.iloc[5, col]).strip()
                esc = "PB" if "PB" in raw_esc else ("Actual" if "FEB" in raw_esc or "ACT" in raw_esc or "PY" in raw_esc else raw_esc)
                if esc == 'nan' or not esc: esc = "Actual"
                
                dum = raw_dummy[:7]
                rango_m = str(df.iloc[8, col]).strip()
                
                dict_cabeceras[col] = {
                    'escenario': esc,
                    'dummy': dum,
                    'rango_meses': rango_m
                }

            # --- PROCESAR SECCIÓN INGREDIENTES ---
            for idx in range(idx_composicion + 1, fin_ingredientes):
                cod_mp = str(df.iloc[idx, 0]).strip()
                nombre_mp = str(df.iloc[idx, 1]).strip()
                
                if cod_mp == 'nan' or not cod_mp or 'TOTAL' in nombre_mp.upper() or 'CÓD' in cod_mp.upper():
                    continue
                    
                for col, meta in dict_cabeceras.items():
                    if col >= df.shape[1]: continue
                    valor_inclusion = df.iloc[idx, col]
                    
                    # Saltar si está vacío, es cero o no es numérico
                    if pd.isna(valor_inclusion) or str(valor_inclusion).strip() == '':
                        continue
                    try:
                        val_float = float(valor_inclusion)
                        if val_float == 0: continue
                    except:
                        continue
                    
                    lista_meses = expandir_meses(meta['rango_meses'])
                    for m in lista_meses:
                        data_ingredientes_total.append({
                            'Escenario': meta['escenario'],
                            'Mes': m,
                            'Cód. Dum': meta['dummy'],
                            'Cód. Mat': cod_mp,
                            'Materia prima': nombre_mp,
                            'Peso (Kilos)': val_float,
                            'Original Fila': meta['rango_meses']
                        })

            # --- PROCESAR SECCIÓN NUTRIENTES ---
            # Identificamos en qué columna empieza el primer valor numérico de Análisis
            # Buscando la fila donde están las palabras 'Valor'
            col_datos_analisis_inicio = 4
            for c in range(2, df.shape[1]):
                if 'VALOR' in str(df.iloc[idx_analisis + 1, c]).upper():
                    col_datos_analisis_inicio = c
                    break

            for idx in range(idx_analisis + 2, df.shape[0]):
                tipo_cara = str(df.iloc[idx, 0]).strip()
                cod_nut = str(df.iloc[idx, 1]).strip()
                nombre_nut = str(df.iloc[idx, 2]).strip()
                unidad = str(df.iloc[idx, 3]).strip()
                
                if cod_nut == 'nan' or not cod_nut or 'CÓD' in cod_nut.upper() or 'TIPO' in tipo_cara.upper():
                    continue
                
                # Recorremos las columnas numéricas de datos de análisis
                # Ojo: la sección análisis puede estar corrida, así que nos alineamos usando el dict_cabeceras maestro
                for col in range(col_datos_analisis_inicio, df.shape[1]):
                    valor_nutriente = df.iloc[idx, col]
                    
                    if pd.isna(valor_nutriente) or str(valor_nutriente).strip() == '':
                        continue
                    try:
                        val_nut_float = float(valor_nutriente)
                    except:
                        continue
                    
                    # Sincronización maestra: Buscamos la cabecera correspondiente en la fila 7
                    # Si hay un desfase de columnas visuales, lo acoplamos buscando el código de columna correcto
                    # En tu Excel, la columna de datos de Nutrientes (col) se corresponde exactamente con la columna (col) si están alineadas,
                    # pero si hay una columna extra de 'Valor', la corregimos dinámicamente:
                    col_maestra = col
                    if col not in dict_cabeceras and (col - 1) in dict_cabeceras:
                        col_maestra = col - 1
                        
                    if col_maestra in dict_cabeceras:
                        meta = dict_cabeceras[col_maestra]
                        lista_meses = expandir_meses(meta['rango_meses'])
                        
                        for m in lista_meses:
                            data_nutrientes_total.append({
                                'Escenario': meta['escenario'],
                                'Mes': m,
                                'Cód. Dum': meta['dummy'],
                                'Tipo': tipo_cara,
                                'Cód. Nutriente': cod_nut,
                                'Característica': nombre_nut,
                                'Unidad': unidad,
                                'Valor Analítico': val_nut_float,
                                'Original Fila': meta['rango_meses']
                            })

        # --- CONSOLIDACIÓN FINAL E INYECCIÓN KARDEX ---
        df_ing_final = pd.DataFrame(data_ingredientes_total)
        df_nut_final = pd.DataFrame(data_nutrientes_total)
        
        if not df_ing_final.empty and k_pb is not None and k_act is not None:
            df_ing_final['Cód. Mat'] = df_ing_final['Cód. Mat'].astype(str).str.strip()
            k_pb['Cod MP'] = k_pb['Cod MP'].astype(str).str.strip()
            k_act['Cod MP'] = k_act['Cod MP'].astype(str).str.strip()
            
            df_ing_final = pd.merge(df_ing_final, k_pb[['Cod MP', 'Mes', 'Precio']], left_on=['Cód. Mat', 'Mes'], right_on=['Cod MP', 'Mes'], how='left')
            df_ing_final = pd.merge(df_ing_final, k_act[['Cod MP', 'Mes', 'Precio']], left_on=['Cód. Mat', 'Mes'], right_on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))
            df_ing_final.drop(columns=['Cod MP_x', 'Cod MP_y'], errors='ignore', inplace=True)

        if not df_ing_final.empty or not df_nut_final.empty:
            st.success("¡Estructura unificada y alineada por Código Maestro exitosamente!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_ing_final.empty:
                    df_ing_final.to_excel(writer, sheet_name="INGREDIENTES", index=False)
                    st.write("### Vista previa final de INGREDIENTES:")
                    st.dataframe(df_ing_final.head(10))
                if not df_nut_final.empty:
                    df_nut_final.to_excel(writer, sheet_name="NUTRIENTES", index=False)
                    st.write("### Vista previa final de NUTRIENTES:")
                    st.dataframe(df_nut_final.head(10))
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel Estructurado y Sincronizado",
                data=output,
                file_name="CONSOLIDADO_MATRIZ_TOTAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error en el proceso de alineación: {e}")
