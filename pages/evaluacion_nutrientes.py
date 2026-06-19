import pandas as pd
import streamlit as st
import io
import re

st.title("Módulo de Separación Avanzada: Ingredientes y Nutrientes")
st.markdown("Carga tu archivo de formulación. El sistema separará automáticamente la información en las pestañas INGREDIENTES y NUTRIENTES.")

def expandir_meses(rango_texto):
    meses_ano = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    meses_map = {m[:3].lower(): m for m in meses_ano}
    
    texto = str(rango_texto).strip()
    if '-' in texto:
        partes = texto.split('-')
        inicio = partes[0].strip()[:3].lower()
        fin = partes[1].strip()[:3].lower()
        
        if inicio in meses_map and fin in meses_map:
            idx_inicio = meses_ano.index(meses_map[inicio])
            idx_fin = meses_ano.index(meses_map[fin])
            if idx_inicio <= idx_fin:
                return meses_ano[idx_inicio:idx_fin+1]
            else:
                return meses_ano[idx_inicio:] + meses_ano[:idx_fin+1]
    
    # Si es un solo mes
    mes_corto = texto[:3].lower()
    if mes_corto in meses_map:
        return [meses_map[mes_corto]]
    
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
            
            # --- LEER CABECERA HORIZONTAL MATRIZ ---
            # Fila 0: Escenario (ej: PY06)
            escenarios = [str(df.iloc[0, c]).strip() for c in range(5, df.shape[1])]
            # Fila 1: Rango de Meses (ej: Ene - Jun)
            rangos_meses = [str(df.iloc[1, c]).strip() for c in range(5, df.shape[1])]
            # Fila 1: Códigos de Fórmulas / Dummies
            dummies = [str(df.iloc[1, c]).strip() for c in range(5, df.shape[1])] # Ajustar índice si cambia de fila
            
            # Intentar buscar dinámicamente las filas de control basadas en texto de la columna B o A
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
            
            # --- PROCESAR SECCIÓN INGREDIENTES (BLOQUE NARANJA) ---
            for idx in range(idx_composicion + 2, fin_ingredientes):
                cod_mp = str(df.iloc[idx, 0]).strip()
                nombre_mp = str(df.iloc[idx, 1]).strip()
                
                if cod_mp == 'nan' or not cod_mp or 'TOTAL' in nombre_mp.upper():
                    continue
                    
                for c_idx, col in enumerate(range(5, df.shape[1])):
                    valor_inclusion = df.iloc[idx, col]
                    if pd.isna(valor_inclusion) or float(valor_inclusion) == 0:
                        continue
                        
                    esc = escenarios[c_idx] if c_idx < len(escenarios) else hoja
                    rango_m = rangos_meses[c_idx] if c_idx < len(rangos_meses) else "Ene - Dic"
                    dum = dummies[c_idx] if c_idx < len(dummies) else ""
                    
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

            # --- PROCESAR SECCIÓN NUTRIENTES (BLOQUE GRIS) ---
            for idx in range(idx_analisis + 2, df.shape[0]):
                tipo_cara = str(df.iloc[idx, 0]).strip()
                cod_nut = str(df.iloc[idx, 1]).strip()
                nombre_nut = str(df.iloc[idx, 2]).strip()
                unidad = str(df.iloc[idx, 3]).strip()
                
                if cod_nut == 'nan' or not cod_nut:
                    continue
                    
                for c_idx, col in enumerate(range(5, df.shape[1])):
                    valor_nutriente = df.iloc[idx, col]
                    if pd.isna(valor_nutriente):
                        continue
                        
                    esc = escenarios[c_idx] if c_idx < len(escenarios) else hoja
                    rango_m = rangos_meses[c_idx] if c_idx < len(rangos_meses) else "Ene - Dic"
                    dum = dummies[c_idx] if c_idx < len(dummies) else ""
                    
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
            st.success("¡Estructura idéntica procesada en dos pestañas!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if not df_ing_final.empty:
                    df_ing_final.to_excel(writer, sheet_name="INGREDIENTES", index=False)
                    st.write("### Pestaña INGREDIENTES Generada:")
                    st.dataframe(df_ing_final.head(5))
                if not df_nut_final.empty:
                    df_nut_final.to_excel(writer, sheet_name="NUTRIENTES", index=False)
                    st.write("### Pestaña NUTRIENTES Generada:")
                    st.dataframe(df_nut_final.head(5))
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel Estructurado",
                data=output,
                file_name="RESULTADO_INGREDIENTES_Y_NUTRIENTES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error procesando la matriz: {e}")
