import pandas as pd
import glob

def generar_reporte_comparativo():
    # 1. Identificar el archivo Excel automáticamente en la carpeta local
    archivos = glob.glob("*.xlsx")
    if not archivos:
        print("Error: No se encontró ningún archivo .xlsx en esta carpeta.")
        return
        
    archivo = archivos[0]
    print(f"Procesando archivo: {archivo}")
    
    # Hojas de fórmulas a procesar
    hojas_targets = {'PB': 'F.PB', 'Ant': 'F.PYAnterior', 'Act': 'F.PYActual'}
    id_cols = ['Dummy', 'Mes', 'SKU', 'Nombre del producto']
    
    data_formulas = []
    data_nutrientes = []
    
    # 2. Procesamiento Dinámico de Columnas (Melt)
    for label, hoja in hojas_targets.items():
        try:
            df = pd.read_excel(archivo, sheet_name=hoja)
            
            # Asegurar que las columnas de identificación existan en la hoja
            cols_presentes = [c for c in id_cols if c in df.columns]
            
            # --- SEPARACIÓN DINÁMICA ---
            # Ingredientes: Todo lo que empiece con 'PC_'
            cols_ingredientes = [c for c in df.columns if str(c).startswith('PC_')]
            # Nutrientes: Todo lo que NO sea ID ni sea ingrediente
            cols_nutrientes = [c for c in df.columns if c not in cols_presentes and c not in cols_ingredientes]
            
            # Trasponer Ingredientes (Receta)
            if cols_ingredientes:
                df_ing_v = df.melt(id_vars=cols_presentes, value_vars=cols_ingredientes, var_name='Cod MP', value_name='Inclusion')
                df_ing_v['Escenario'] = label
                data_formulas.append(df_ing_v)
                
            # Trasponer Nutrientes
            if cols_nutrientes:
                df_nut_v = df.melt(id_vars=cols_presentes, value_vars=cols_nutrientes, var_name='Cod Nutriente', value_name='Valor Nutriente')
                df_nut_v['Escenario'] = label
                data_nutrientes.append(df_nut_v)
                
        except Exception as e:
            print(f"Aviso: No se pudo procesar la hoja {hoja}. Error: {e}")

    # ==========================================
    # SECCIÓN INGREDIENTES + PRECIOS (Tu base)
    # ==========================================
    if data_formulas:
        master_f = pd.concat(data_formulas).dropna(subset=['Inclusion'])
        
        # Carga de Kardex para cruce de precios
        try:
            k_pb = pd.read_excel(archivo, sheet_name='Kardex_PB')
            k_act = pd.read_excel(archivo, sheet_name='Kardex_Actual')
            
            # Asegurar formato texto para códigos evitando problemas de formato científico
            master_f['Cod MP'] = master_f['Cod MP'].astype(str).str.strip()
            k_pb['Cod MP'] = k_pb['Cod MP'].astype(str).str.strip()
            k_act['Cod MP'] = k_act['Cod MP'].astype(str).str.strip()
            
            # Cruce de precios (Kardex)
            df_costos = pd.merge(master_f, k_pb[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left')
            df_costos = pd.merge(df_costos, k_act[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))
            
            # Cálculos de costos e indicadores
            df_costos['Costo_con_KPB'] = df_costos['Inclusion'] * df_costos['Precio_KPB']
            df_costos['Costo_con_KACT'] = df_costos['Inclusion'] * df_costos['Precio_KACT']
            df_costos['ALERTA'] = df_costos.apply(lambda x: 'FALTA PRECIO' if pd.isna(x['Precio_KACT']) else 'OK', axis=1)
            
            # Pivot resumen de recetas (Actual vs PB)
            resumen_recetas = df_costos.pivot_table(index=['Dummy', 'Mes'], columns='Escenario', values='Costo_con_KACT', aggfunc='sum')
            if 'Act' in resumen_recetas.columns and 'PB' in resumen_recetas.columns:
                resumen_recetas['Diff_Act_vs_PB'] = resumen_recetas['Act'] - resumen_recetas['PB']
            else:
                resumen_recetas = pd.DataFrame()
        except Exception as e:
            print(f"Error al procesar costos con Kardex: {e}")
            df_costos = master_f
            resumen_recetas = pd.DataFrame()
    else:
        df_costos = pd.DataFrame()
        resumen_recetas = pd.DataFrame()

    # ==========================================
    # SECCIÓN NUTRIENTES (Nueva Pestaña)
    # ==========================================
    if data_nutrientes:
        master_nutrientes = pd.concat(data_nutrientes).dropna(subset=['Valor Nutriente'])
    else:
        master_nutrientes = pd.DataFrame()

    # ==========================================
    # EXPORTACIÓN FINAL A EXCEL
    # ==========================================
    archivo_salida = "RESULTADO_COMPARATIVO_CON_NUTRIENTES.xlsx"
    with pd.ExcelWriter(archivo_salida) as writer:
        if not df_costos.empty:
            # Aquí ya van incluidos los campos Precio_KPB y Precio_KACT que solicitaste
            df_costos.to_excel(writer, sheet_name="DATA_PARA_POWERBI", index=False)
        if not resumen_recetas.empty:
            resumen_recetas.to_excel(writer, sheet_name="RESUMEN_RECETAS")
        if not df_costos.empty and 'ALERTA' in df_costos.columns:
            df_costos[df_costos['ALERTA'] == 'FALTA PRECIO'].to_excel(writer, sheet_name="ALARMAS", index=False)
        if not master_nutrientes.empty:
            # Pestaña independiente para los nutrientes transpuestos
            master_nutrientes.to_excel(writer, sheet_name="VERTICAL_NUTRIENTES", index=False)
            
    print(f"¡Proceso completado con éxito! Archivo generado: {archivo_salida}")

if __name__ == "__main__":
    generar_reporte_comparativo()
