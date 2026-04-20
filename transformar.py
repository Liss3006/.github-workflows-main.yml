import pandas as pd
import glob

def generar_reporte_comparativo():
    archivo = glob.glob("*.xlsx")[0]
    
    # 1. Carga de datos con tipos de datos forzados (evita el error E+16)
    dict_hojas = {
        'PB': 'F.PB', 'Anterior': 'F.PYAnterior', 'Actual': 'F.PYActual',
        'K_PB': 'Kardex_PB', 'K_Actual': 'Kardex_Actual'
    }
    
    # Procesar fórmulas a vertical
    data_formulas = []
    for label, hoja in {'PB': 'F.PB', 'Ant': 'F.PYAnterior', 'Act': 'F.PYActual'}.items():
        df = pd.read_excel(archivo, sheet_name=hoja)
        # Seleccionamos columnas de identificación y de ingredientes
        id_cols = ['Dummy', 'Mes', 'SKU', 'Nombre del producto']
        df_v = df.melt(id_vars=id_cols, var_name='Cod MP', value_name='Inclusion')
        df_v['Escenario'] = label
        data_formulas.append(df_v)
    
    master_f = pd.concat(data_formulas).dropna(subset=['Inclusion'])

    # 2. Cruce con Kardex (PB y Actual)
    k_pb = pd.read_excel(archivo, sheet_name='Kardex_PB')
    k_act = pd.read_excel(archivo, sheet_name='Kardex_Actual')

    # Unimos todo en una gran tabla de costos
    df_costos = pd.merge(master_f, k_pb[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left')
    df_costos = pd.merge(df_costos, k_act[['Cod MP', 'Mes', 'Precio']], on=['Cod MP', 'Mes'], how='left', suffixes=('_KPB', '_KACT'))

    # 3. Cálculos de Diferenciales (Lo que pides en tus tablas)
    df_costos['Costo_con_KPB'] = df_costos['Inclusion'] * df_costos['Precio_KPB']
    df_costos['Costo_con_KACT'] = df_costos['Inclusion'] * df_costos['Precio_KACT']
    
    # Alarma de precios faltantes
    df_costos['ALERTA'] = df_costos.apply(lambda x: 'FALTA PRECIO' if pd.isna(x['Precio_KACT']) else 'OK', axis=1)

    # 4. Generación de Resúmenes Estilo Tabla (Pivots)
    # Tabla: Diferencial Formulas (Actual vs PB) usando Kardex Actual
    resumen_recetas = df_costos.pivot_table(
        index=['Dummy', 'Mes'], 
        columns='Escenario', 
        values='Costo_con_KACT', 
        aggfunc='sum'
    )
    resumen_recetas['Diff_Act_vs_PB'] = resumen_recetas['Act'] - resumen_recetas['PB']

    # Exportar resultados
    with pd.ExcelWriter("RESULTADO_COMPARATIVO.xlsx") as writer:
        df_costos.to_excel(writer, sheet_name="DATA_PARA_POWERBI", index=False)
        resumen_recetas.to_excel(writer, sheet_name="RESUMEN_RECETAS")
        df_costos[df_costos['ALERTA'] == 'FALTA PRECIO'].to_excel(writer, sheet_name="ALARMAS")

if __name__ == "__main__":
    generar_reporte_comparativo()
