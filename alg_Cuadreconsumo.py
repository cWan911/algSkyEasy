import pandas as pd
import openpyxl

def obtener_tipo_canje(respuesta):
    if respuesta == 4:
        return "Pajapita"
    elif respuesta == 103:
        return "Vale Puma"
    elif respuesta == 10:
        return "Depositos"
    elif respuesta == 8:
        return "Nicaragua"
    else:
        return "Otro Canje"

def agregar_guion(placa):
    return placa[0] + '-' + placa[1:]

try:
    archivo_excel_1 = 'C:/Users/bi/Documents/Operiaciones_TLC/dashVale/hstvale.xlsx'
    archivo_excel_2 = 'C:/Users/bi/Documents/Operiaciones_TLC/dashVale/hstRuta.xlsx'
    # Leer las hojas
    df_vale = pd.read_excel(archivo_excel_1, sheet_name='dtVale')
    df_sky = pd.read_excel(archivo_excel_1, sheet_name='dtSky')
    df_ruta = pd.read_excel(archivo_excel_2, sheet_name='hjRuta')
    
#    print("\nPrimeras filas de dtSky:")
#    print(df_sky.head())
except FileNotFoundError:
    print("El archivo Excel no fue encontrado.")
except Exception as e:
    print("Ocurrió un error al leer el archivo Excel:", e)

#Agregar columnas TipoVale y TipoCanje
df_tpVale = df_ruta[['Id_R','tpVale']]
df_vale['TipoCanje'] = df_vale['Resp.'].apply(obtener_tipo_canje)
df_resultado = pd.merge(df_vale, df_tpVale, left_on='Destino', right_on='Id_R', how='left')
df_resultado.drop(columns=['Id_R'], inplace=True)
df_resultado.rename(columns={'tpVale': 'TipoVale'}, inplace=True)
#Agregar guion a placas skyfleet
df_sky['Placa'] = df_sky['Placa'].apply(lambda x: agregar_guion(x))

# Convertir la columna de fecha de texto a tipo datetime
df_sky['Fecha Venta'] = df_sky['Fecha Venta'].str.replace('.', '')
df_sky['FechaFor'] = pd.to_datetime(df_sky['Fecha Venta'], format='%d/%m/%Y %I:%M:%S %p')
# Agregar una columna con la fecha en formato 'dd/mm/yyyy'
df_sky['FechaNor'] = df_sky['FechaFor'].dt.strftime('%d/%m/%Y')
df_sky['FechaNor'] =pd.to_datetime(df_sky['FechaNor'], format='%d/%m/%Y')
# Agregar una columna con la fecha y hora en el formato deseado
df_sky['FechaCom'] = pd.to_datetime(df_sky['FechaFor'], format='%d/%m/%Y %H:%M:%S')

# Convertir fechas a formato datetime
df_vale['Fecha'] = pd.to_datetime(df_vale['Fecha'], format='%d/%m/%Y')
#filtroVal = (df_vale['Fecha'] >= '2024-02-01') & (df_vale['Resp.'] == 4)
filtroVal = df_vale['Resp.'] == 4
df_vales = df_vale[filtroVal]
df_vales.sort_values(by='Fecha', ascending=False)
df_vales['IdVales'] = df_vales['Vale'].apply(lambda x: int(x[2:]))
df_vales['Consumo']=None
# Ordenar la tabla "sky" por fechaDespacho de forma descendente
df_sky.sort_values(by='FechaCom', ascending=False)
          
for index, row in df_sky.iterrows():
    placa = row['Placa']
    galones_DesT = row['Volumen']
    galones_sky = row['Volumen']
    fecha_filter_sky = row['FechaNor']
    fecha_despacho_sky = row['FechaCom']
    id_venta_sky = row['Id Venta']
    
    if galones_DesT < 1:
        break

    # Filtrar vales por placa y consumo diferente a si
    vales_filtrados = df_vales[(df_vales['Placa'] == placa) & (df_vales['Consumo']!= 'Si')]
    # Filtrar sky para saber si hay mas de un 1 registro con la misma placa en la misma fecha
    sky_filtrados =df_sky[(df_sky['Placa']==placa) & (df_sky['FechaNor']==fecha_filter_sky)]
    nvalores = sky_filtrados.shape[0]
    
    if nvalores > 1:
        vales_filtrados = vales_filtrados.sort_values(by='IdVales', ascending=False)

    # Realizar otro filtro, con la condicion si la fecha de sky esta en los datos filtrados de vales que corresponde a la placa 
    # si no esta que filtre a la fecha proxima
    if fecha_filter_sky in vales_filtrados[vales_filtrados['Fecha'] == fecha_filter_sky]:
        vales_filtrados = vales_filtrados[vales_filtrados['Fecha'] == fecha_filter_sky]
    else:
        vales_filtrados = vales_filtrados[vales_filtrados['Fecha'] < fecha_filter_sky]
        vales_filtrados = vales_filtrados.iloc[:1]         
    # si sky tiene mas de 1 un dato, ordenar los vales de mayor a menor
        
    vales_filtrados['Id Venta'] = id_venta_sky 
    galones_acumulados = 0
    # Iterar sobre los vales filtrados
    for _, vale_row in vales_filtrados.iterrows():
        galones_vale = vale_row['Gal']

        # Asegurarse de no exceder los galones de sky
        if galones_sky > 0:
            galones_consumidos = min(galones_sky, galones_vale)
            galones_sky -= galones_consumidos

            # Agregar nueva información a una fila específica por su idVale
            df_vales.loc[df_vales['Vale'] == vale_row['Vale'] ,['Consumo','idVenta', 'fechaDespacho', 'galones_sky','galones_Consumo','galones_TotalDespacho','FechaFiltro','nval','order']] = ["Si",id_venta_sky, fecha_despacho_sky, galones_sky,galones_consumidos,galones_DesT,fecha_filter_sky,nvalores,vale_row['IdVales']]


# Guardar el resultado en un archivo CSV
df_vales.to_csv('C:/Users/bi/Documents/Operiaciones_TLC/dashVale/resultado17.csv', index=False)
print("Documento guardado exitosamente")