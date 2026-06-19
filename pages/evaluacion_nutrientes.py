import pandas as pd
import streamlit as st
import io

st.title("Módulo de Separación Avanzada: Ingredientes y Nutrientes")
st.markdown("Carga tu archivo de formulación. El sistema separará automáticamente la información en las pestañas INGREDIENTES y NUTRIENTES.")

def expandir_meses(rango_texto):
    meses_ano = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    meses_map = {m[:3].lower(): m for m in meses_ano}
    
    texto = str(rango_texto).strip()
    
    # Si viene con la palabra "PB" o texto extra, limpiarla para detectar el mes
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
                
    # Intentar detectar un mes único
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
                
            # Leer el archivo sin cabecera para controlar las filas por su índice real de Excel
            df = pd.read_excel(archivo_cargado, sheet_name=hoja, header=None)
            
            # --- COORDINADAS CORREGIDAS SEGÚN LA MATRIZ REAL ---
            # Excel Fila 6 (Índice 5): Estab / Escenario (ej: PY11 PB26)
            fila_escenarios = df.iloc[5, :]
            # Excel Fila 7 (Índice 6): Cód. Dummy (ej: HD144121 -> queremos primeros 7: HD14412)
            fila_dummies = df.iloc[6, :]
            # Excel Fila 9 (Índice 8): Carpeta / Meses (ej: Ene - Abr 2026 PB)
            fila_carpetas = df.iloc[8, :]
            
            # Buscar dinámicamente dónde empiezan las secciones Composición y Análisis
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
            
            # --- DETECTAR COLUMNAS CON DATOS (De la columna E en adelante) ---
            columnas_datos = range(4, df.shape[1])
            
            # --- PROCESAR SECCIÓN INGREDIENTES ---
            for idx in range(idx_composicion + 1, fin_ingredientes):
                cod_mp = str(df.iloc[idx, 0]).strip()
                nombre_mp = str(df.iloc[idx, 1]).strip()
                
                if cod_mp == 'nan' or not cod_mp or 'TOTAL' in nombre_mp.upper() or 'CÓD' in cod_mp.upper():
                    continue
                    
                for col in columnas_datos:
                    if col >= df.shape[1]: continue
                    valor_inclusion = df.iloc[idx, col]
                    
                    if pd.isna(valor_inclusion) or str(valor_inclusion).strip() == '' or float(valor_inclusion) == 0:
                        continue
                        
                    # Extraer y limpiar metadatos de las filas superiores corregidas
                    raw_esc = str(fila_escenarios.iloc[col]).strip()
                    esc = "PB" if "PB" in raw_esc else ("Actual" if "FEB" in raw_esc or "ACT" in raw_esc else raw_esc)
                    
                    # Cód Dummy: Tomar solo los primeros 7 caracteres de la celda de la Fila 7
                    raw_dummy = str(fila_dummies.iloc[col]).strip()
                    dum = raw_dummy[:7] if raw_dummy != 'nan' else ""
                    
                    rango_m = str(fila_carpetas.iloc[col]).strip()
                    
                    lista_meses = expandir_meses(rango_m)
                    for m in lista_meses:
                        data_ingredientes_total.append({
                            'Escenario': esc,
                            'Mes': m,
                            'Cód. Dum': dum,
                            'Cód. Mat': cod_mp,
                            'Materia prima': nombre_mp,
                            'Peso (Kilos)': float(valor_inclusion),
                            'Original Fila': rango_m
                        })

            # --- PROCESAR SECCIÓN NUTRIENTES ---
            for idx in range(idx_analisis + 1, df.shape[0]):
                tipo_cara = str(df.iloc[idx, 0]).strip()
                cod_nut = str(df.iloc[idx, 1]).strip()
                nombre_nut = str(df.iloc[idx, 2]).strip()
                unidad = str(df.iloc[idx, 3]).strip()
                
                if cod_nut == 'nan' or not cod_nut or 'CÓD' in cod_nut.upper() or 'TIPO' in tipo_cara.upper():
                    continue
                    
                for col in columnas_datos:
                    if col >= df.shape[1]: continue
                    valor_nutriente = df.iloc[idx, col]
                    
                    if pd.isna(valor_nutriente) or str(valor_nutriente).strip() == '':
                        continue
                        
                    raw_esc = str(fila_escenarios.iloc[col]).strip()
                    esc = "PB" if "PB" in raw_esc else ("Actual" if "FEB" in raw_esc or "ACT" in raw_esc else raw_esc)
                    
                    raw_dummy = str(fila_dummies.iloc[col]).strip()
                    dum = raw_dummy[:7] if raw_dummy != 'nan' else ""
                    
                    rango_m = str(fila_carpetas.iloc[col]).strip()
                    
                    lista_meses = expandir_meses(rango_m)
                    for m in lista_meses:
                        data_nutrientes_total.append({
                            'Escenario': esc,
                            'Mes': m,
                            'Cód. Dum': dum,
                            'Tipo': tipo_cara,
                            'Cód. Nutriente': cod_nut,
                            'Característica': nombre_nut,
                            'Unidad': unidad,
                            'Valor Analítico': float(valor_nutriente),
                            'Original Fila': rango_m
                        })

        # --- GENERAR DF E INYECTAR COSTOS KARDEX ---
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
            st.success("¡Estructura mapeada con cabeceras correctas y corte de Dummy a 7 dígitos!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_ing_final.empty:
                    df_ing_final.to_excel(writer, sheet_name="INGREDIENTES", index=False)
                    st.write("### Vista previa de INGREDIENTES:")
                    st.dataframe(df_ing_final.head(5))
                if not df_nut_final.empty:
                    df_nut_final.to_excel(writer, sheet_name="NUTRIENTES", index=False)
                    st.write("### Vista previa de NUTRIENTES:")
                    st.dataframe(df_nut_final.head(5))
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel Estructurado Completo",
                data=output,
                file_name="RESULTADO_INGREDIENTES_Y_NUTRIENTES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error procesando la matriz: {e}")
