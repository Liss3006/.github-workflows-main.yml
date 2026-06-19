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
            
            # --- LOCALIZACIÓN DE BLOQUES DE CONTROL ---
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
            
            # --- CAPTURA ASOCIATIVA DE CABECERAS (Desde columna E / Índice 4 en adelante) ---
            # Guardamos un diccionario mapeando cada columna numérica de Composición con su metadata exacta
            mapa_cabeceras = {}
            
            # Fila 6 (Índice 5): Escenarios
            # Fila 7 (Índice 6): Dummies (recorte a 7 dígitos)
            # Fila 9 (Índice 8): Carpetas/Meses
            for col in range(4, df.shape[1]):
                raw_esc = str(df.iloc[5, col]).strip()
                esc = "PB" if "PB" in raw_esc else ("Actual" if "FEB" in raw_esc or "ACT" in raw_esc else raw_esc)
                
                raw_dummy = str(df.iloc[6, col]).strip()
                dum = raw_dummy[:7] if raw_dummy != 'nan' else ""
                
                rango_m = str(df.iloc[8, col]).strip()
                
                # Almacenamos la estructura limpia indexada por su posición relativa
                mapa_cabeceras[col] = {
                    'escenario': esc,
                    'dummy': dum,
                    'rango_meses': rango_m
                }

            # --- PROCESAR SECCIÓN INGREDIENTES (Datos en Columna E -> Índice 4) ---
            for idx in range(idx_composicion + 1, fin_ingredientes):
                cod_mp = str(df.iloc[idx, 0]).strip()
                nombre_mp = str(df.iloc[idx, 1]).strip()
                
                if cod_mp == 'nan' or not cod_mp or 'TOTAL' in nombre_mp.upper() or 'CÓD' in cod_mp.upper():
                    continue
                    
                for col in range(4, df.shape[1]):
                    valor_inclusion = df.iloc[idx, col]
                    if pd.isna(valor_inclusion) or str(valor_inclusion).strip() == '' or float(valor_inclusion) == 0:
                        continue
                    
                    meta = mapa_cabeceras[col]
                    lista_meses = expandir_meses(meta['rango_meses'])
                    
                    for m in lista_meses:
                        data_ingredientes_total.append({
                            'Escenario': meta['escenario'],
                            'Mes': m,
                            'Cód. Dum': meta['dummy'],
                            'Cód. Mat': cod_mp,
                            'Materia prima': nombre_mp,
                            'Peso (Kilos)': float(valor_inclusion),
                            'Original Fila': meta['rango_meses']
                        })

            # --- PROCESAR SECCIÓN NUTRIENTES (Datos en Columna F -> Índice 5) ---
            # El primer valor numérico de Análisis está desplazado +1 columna a la derecha
            for idx in range(idx_analisis + 1, df.shape[0]):
                tipo_cara = str(df.iloc[idx, 0]).strip()
                cod_nut = str(df.iloc[idx, 1]).strip()
                nombre_nut = str(df.iloc[idx, 2]).strip()
                unidad = str(df.iloc[idx, 3]).strip()
                
                if cod_nut == 'nan' or not cod_nut or 'CÓD' in cod_nut.upper() or 'TIPO' in tipo_cara.upper():
                    continue
                    
                for col in range(5, df.shape[1]):
                    valor_nutriente = df.iloc[idx, col]
                    if pd.isna(valor_nutriente) or str(valor_nutriente).strip() == '':
                        continue
                    
                    # Sincronizamos restando 1 al índice de columna para acoplarnos al mapa de Composición
                    col_cabecera_sincro = col - 1
                    if col_cabecera_sincro in mapa_cabeceras:
                        meta = mapa_cabeceras[col_cabecera_sincro]
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
                                'Valor Analítico': float(valor_nutriente),
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
            st.success("¡Sincronización de cabeceras completada exitosamente!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_ing_final.empty:
                    df_ing_final.to_excel(writer, sheet_name="INGREDIENTES", index=False)
                    st.write("### Vista previa de INGREDIENTES Sincronizados:")
                    st.dataframe(df_ing_final.head(5))
                if not df_nut_final.empty:
                    df_nut_final.to_excel(writer, sheet_name="NUTRIENTES", index=False)
                    st.write("### Vista previa de NUTRIENTES Sincronizados:")
                    st.dataframe(df_nut_final.head(5))
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel Estructurado y Sincronizado",
                data=output,
                file_name="CONSOLIDADO_MATRIZ_TOTAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error en el proceso de alineación: {e}")
