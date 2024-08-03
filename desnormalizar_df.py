import pandas as pd

# Leer datos desde un archivo .xlsx
file_path = 'ruta_a_tu_archivo.xlsx'  # Reemplaza con la ruta a tu archivo
df = pd.read_excel(file_path)

# Usar melt para convertir de formato ancho a largo
melted_df = pd.melt(df, id_vars=['Código'], var_name='Variable', value_name='Valor')

# Extraer el mes y el tipo de datos (Cantidad/Venta)
melted_df[['Tipo', 'Mes']] = melted_df['Variable'].str.extract(r'(Cantidad|Venta) (.+)')

# Pivotar para tener columnas separadas para 'Cantidad Comprada' y 'Venta'
normalized_df = melted_df.pivot_table(index=['Código', 'Mes'], columns='Tipo', values='Valor').reset_index()

# Renombrar columnas para mayor claridad
normalized_df.columns.name = None
normalized_df.rename(columns={'Cantidad': 'Cantidad Comprada', 'Venta': 'Venta'}, inplace=True)

# Guardar el DataFrame normalizado a un nuevo archivo .xlsx
output_file_path = 'ruta_a_tu_archivo_normalizado.xlsx'  # Reemplaza con la ruta de salida deseada
normalized_df.to_excel(output_file_path, index=False)

# Mostrar el DataFrame normalizado
print(normalized_df)
