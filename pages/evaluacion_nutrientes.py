import pandas as pd
import streamlit as st
import io

st.title("Evaluación de Nutrientes Dinámica - Balanceado")
st.write("Carga tu archivo de fórmulas horizontal para trasponer y separar Recetas y Nutrientes de forma vertical.")

# Carga de archivo en la interfaz web de Streamlit
archivo_cargado = st.file_uploader("Elige tu archivo de Excel (.xlsx)", type=["xlsx"])

if archivo_cargado is not None:
    try:
        # Hojas de fórmulas objetivo
        hojas_formulas = {'PB': 'F.PB', 'Ant': 'F.PYAnterior', 'Act': 'F.PYActual'}
        
        data_recetas_total = []
        data_nutrientes_total = []
        
        # Leer el archivo de Excel
        xl = pd.ExcelFile(archivo_cargado)
        
        for label, hoja in hojas_formulas.items():
            if hoja in xl.sheet_names:
                # Leer la hoja completa
                df = pd.read_excel(archivo_cargado, sheet_name=hoja)
                
                # 1. Identificar columnas de cabecera (Azul): Todo lo que NO sea ingrediente (PC_) ni nutriente
                # Asumimos que los ingredientes empiezan con 'PC_' y los nutrientes son el resto a la derecha.
                # Definimos explícitamente las columnas fijas de identificación que se deben repetir:
                id_cols = [c for c in df.columns if str(c).strip() in ['Dummy', 'Mes', 'SKU', 'Nombre del producto', 'Costo Lote', 'Kg Lote']]
                
                # Si por variaciones de nombre no encuentra las exactas, toma las primeras 4 columnas como cabecera por defecto
                if not id_cols:
                    id_cols = list(df.columns[:4])
                
                # 2. Separar dinámicamente columnas de Ingredientes (Naranja) y Nutrientes (Gris)
                cols_ingredientes = [c for c in df.columns if str(c).startswith('PC_')]
                cols_nutrientes = [c for c in df.columns if c not in id_cols and c not in cols_ingredientes]
                
                # 3. Trasponer (Melt) bloque de Recetas / Ingredientes (Naranja)
                if cols_ingredientes:
                    df_receta_v = df.melt(
                        id_vars=id_cols,
                        value_vars=cols_ingredientes,
                        var_name='Cod MP',
                        value_name='Inclusion'
                    )
                    df_receta_v['Escenario'] = label
                    data_recetas_total.append(df_receta_v)
                
                # 4. Trasponer (Melt) bloque de Perfil Nutricional (Gris)
                if cols_nutrientes:
                    df_nut_v = df.melt(
                        id_vars=id_cols,
                        value_vars=cols_nutrientes,
                        var_name='Cod Nutriente',
                        value_name='Valor Nutriente'
                    )
                    df_nut_v['Escenario'] = label
                    data_nutrientes_total.append(df_nut_v)
        
        # Consolidar y generar archivo de descarga
        if data_recetas_total or data_nutrientes_total:
            st.success("¡Estructura horizontal traspuesta a tablas verticales con éxito!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Pestaña 1: Recetas (Inclusiones)
                if data_recetas_total:
                    master_recetas = pd.concat(data_recetas_total).dropna(subset=['Inclusion'])
                    master_recetas.to_excel(writer, sheet_name="VERTICAL_RECETAS", index=False)
                    st.write("### Vista previa de Recetas Verticales:")
                    st.dataframe(master_recetas.head(5))
                
                # Pestaña 2: Perfil Nutricional
                if data_nutrientes_total:
                    master_nutrientes = pd.concat(data_nutrientes_total).dropna(subset=['Valor Nutriente'])
                    master_nutrientes.to_excel(writer, sheet_name="VERTICAL_NUTRIENTES", index=False)
                    st.write("### Vista previa de Perfil Nutricional Vertical:")
                    st.dataframe(master_nutrientes.head(5))
                    
            output.seek(0)
            
            # Botón de descarga único con las dos pestañas limpias
            st.download_button(
                label="📥 Descargar Resultado Consolidado (.xlsx)",
                data=output,
                file_name="RESULTADO_VERTICAL_CONSOLIDADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se pudo procesar ninguna información de las hojas F.PB, F.PYAnterior o F.PYActual.")
            
    except Exception as e:
        st.error(f"Error crítico en el proceso de trasposición: {e}")
