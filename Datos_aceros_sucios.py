import pandas as pd
import numpy as np

todas_hojas = pd.read_excel('DatasetExcel_Caso Costos.xlsx', sheet_name=None)

#Vamos a verificar que no hayan nulos, cuantas columnas tienen espacios y si hay duplicados
for name, df in todas_hojas.items():
    print(name, df.shape, df.isnull().sum().sum(), "nulos")
    print("Cols con espacios:", [c for c in df.columns if c != c.strip()])
    print("Dupes:", df.duplicated().sum())

#Limpiamos las columnas de espacios para poder usarlas despues y que sean iguales
for nombre_hoja, df in todas_hojas.items():
    df.columns = df.columns.str.strip()
    todas_hojas[nombre_hoja] = df
    print(f"✓ Hoja '{nombre_hoja}' limpiada")

#Creamos unos objetos que tengan las hojas de forma individual para poder manipularla despues.

productos = todas_hojas['Productos']
ordenes = todas_hojas['Órdenes de Producción']
consumo_mp = todas_hojas['Consumo MP']
mano_obra = todas_hojas['ManoObra']
costo_estandar = todas_hojas['CostoEstandar']
cif = todas_hojas['CIF']
hoja7 = todas_hojas['Hoja7']

# Vamos a estandarizar los IDs para poder unir en un solo df
#Vamos a hacerlo de forma individual
#Con la siguiente regla:
# Producto_ID : Producto, Orden produccion, costo estandar,
#Orden_ID: orden produccion, consumo mp, mano obra, cif

productos['ProductoID'] = productos['ProductoID'].str.strip()
ordenes['OrdenID'] = ordenes['OrdenID'].str.strip()
ordenes['ProductoID'] = ordenes['ProductoID'].str.strip()
consumo_mp['OrdenID'] = consumo_mp['OrdenID'].str.strip()
consumo_mp['ConsumoID'] = consumo_mp['ConsumoID'].str.strip()
costo_estandar['ProductoID'] = costo_estandar['ProductoID'].str.strip()
mano_obra['OrdenID'] = mano_obra['OrdenID'].str.strip()
cif['OrdenID'] = cif['OrdenID'].str.strip()

#Vamos a merger en un solo dt

df_completo = ordenes.merge(
    productos,
    on='ProductoID',
    how='left'
)
df_con_consumo = df_completo.merge(
    consumo_mp,
    on='OrdenID',
    how='left'
)
df_con_mano = df_con_consumo.merge(
    mano_obra,
    on='OrdenID',
    how='left'
)
df_detallado = df_con_mano.merge(
    cif,
    on='OrdenID',
    how='left'
)
df_final = df_detallado.merge(
    costo_estandar,
    on='ProductoID',
    how='left'
)
print(f"\n TABLA FINAL: {len(df_final)} filas × {len(df_final.columns)} columnas")

#Vemos todas las columnas:
print("Columnas en tu DataFrame final:")
for i, col in enumerate(df_final.columns, 1):
    print(f"{i}. {col}")

print(df_final.head())
#Ahora generamos columnas extra:

df_powerbi = df_final.copy()
df_powerbi = df_powerbi.rename(columns={
    'CostoTotal_x': 'CostoMP_Real',
    'CostoTotal_y': 'CostoMOD_Real',
    'Costo': 'CostoCIF_Real'
})
#Sacamos el costo real total que s la suma de todos los costos reales
df_powerbi['CostoTotal_Real'] = (
    df_powerbi['CostoMP_Real'].fillna(0) +
    df_powerbi['CostoMOD_Real'].fillna(0) +
    df_powerbi['CostoCIF_Real'].fillna(0)
)
#Sacamos el costo mp real unitario  que s la division entre el costo mp real entre la cantidad producida
df_powerbi['CostoMP_RealUnit'] = np.where(
    df_powerbi['CantidadProducida'] > 0,
    df_powerbi['CostoMP_Real'] / df_powerbi['CantidadProducida'],
    0
)
#SacamosCosto MOD_RealUnit
df_powerbi['CostoMOD_RealUnit'] = np.where(
    df_powerbi['CantidadProducida'] > 0,
    df_powerbi['CostoMOD_Real'] / df_powerbi['CantidadProducida'],
    0
)
#SacamosCosto CIF_RealUnit
df_powerbi['CostoCIF_RealUnit'] = np.where(
    df_powerbi['CantidadProducida'] > 0,
    df_powerbi['CostoCIF_Real'] / df_powerbi['CantidadProducida'],
    0
)
#Sacamos costo total por unidad
df_powerbi['CostoTotal_RealUnit'] = np.where(
    df_powerbi['CantidadProducida'] > 0,
    df_powerbi['CostoTotal_Real'] / df_powerbi['CantidadProducida'],
    0
)

#Sacamos las varianzas para poder comparar con lo que sacamos real
#En orden son :cariación mp,mod,cif, la total. QUe es simplemente la real menos la standar. Si es positiva significa que hay problemas
df_powerbi['VarMP_Unit'] = df_powerbi['CostoMP_RealUnit'] - df_powerbi['CostoMP_Std']
df_powerbi['VarMOD_Unit'] = df_powerbi['CostoMOD_RealUnit'] - df_powerbi['CostoMOD_Std']
df_powerbi['VarCIF_Unit'] = df_powerbi['CostoCIF_RealUnit'] - df_powerbi['CostoCIF_Std']
df_powerbi['VarTotal_Unit'] = df_powerbi['CostoTotal_RealUnit'] - df_powerbi['CostoTotalStd']


#Sacamos los totales, ingresos totales, utilidad bruta y el amrgen porcentual
df_powerbi['Ingresos'] = df_powerbi['CantidadProducida'] * df_powerbi['PrecioVentaUnit']
df_powerbi['UtilidadBruta'] = df_powerbi['Ingresos'] - df_powerbi['CostoTotal_Real']
df_powerbi['MargenBruto_Pct'] = np.where(
    df_powerbi['Ingresos'] > 0,
    (df_powerbi['UtilidadBruta'] / df_powerbi['Ingresos']) * 100,
    0
)
#Sacamos un estado de varianza, que estandariza que
df_powerbi['Estado_Varianza'] = np.where(
    df_powerbi['VarTotal_Unit'] < 0, 'Eficiente',
    np.where(df_powerbi['VarTotal_Unit'] <= 1, 'Alerta', 'Critico')
)
#Armamos bien el formato de fechas
df_powerbi['Fecha'] = pd.to_datetime(df_powerbi['Fecha'])
df_powerbi['Mes'] = df_powerbi['Fecha'].dt.strftime('%Y-%m')
df_powerbi['NombreMes'] = df_powerbi['Fecha'].dt.strftime('%B %Y')

#Ordenamos las columnas

columnas_ordenadas = [
    'OrdenID', 'Fecha', 'ProductoID', 'CantidadProducida', 'Planta',
    'NombreProducto', 'TipoProducto', 'PesoKg', 'PrecioVentaUnit',
    'CostoMP_Std', 'CostoMOD_Std', 'CostoCIF_Std', 'CostoTotalStd',
    'CostoMP_Real', 'CostoMOD_Real', 'CostoCIF_Real', 'CostoTotal_Real',
    'CostoMP_RealUnit', 'CostoMOD_RealUnit', 'CostoCIF_RealUnit', 'CostoTotal_RealUnit',
    'VarMP_Unit', 'VarMOD_Unit', 'VarCIF_Unit', 'VarTotal_Unit',
    'Ingresos', 'UtilidadBruta', 'MargenBruto_Pct', 'Estado_Varianza', 'Mes', 'NombreMes'
]

df_powerbi = df_powerbi[columnas_ordenadas]

#Redondeamos a 2 decimales:
columnas_numericas = df_powerbi.select_dtypes(include=[np.number]).columns
for col in columnas_numericas:
    df_powerbi[col] = df_powerbi[col].round(2)

df_powerbi.to_csv('costos_para_bi.csv',
                  index=False,
                  encoding='utf-8-sig',
                  float_format='%.2f')

print("Archivo exportado exitosamente")