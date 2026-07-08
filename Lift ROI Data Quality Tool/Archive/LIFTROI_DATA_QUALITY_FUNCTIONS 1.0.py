"""
LIFTROI DATA QUALITY FUNCTIONS v1.0 

Developed by: Fernando Monrroy fernando.monrroy@kantar.com

This is a custom script for creating data checks for LiftROI media inputs. AdHoc to the analysis we perform on the methodology.  
"""


import pandas as pd
import numpy as np
import re

def add_change_delta(crosstab_df):
    
    _crosstab = crosstab_df.copy()
    _delta_crosstab = (_crosstab.T.diff(1) / _crosstab.T.shift(1)).T.round(4)
    
    _result_df = pd.merge(
        left=_crosstab,
        right=_delta_crosstab,
        left_index=True,
        right_index=True,
        suffixes=(None, ' change')
    )

    return _result_df


def check_integrity(dataframe, filter_like):
    """
    Funcion que checa la integridad de los datos de actividad para cada medio.
    Si hay multiples métricas de actividad itera por cada una y lo contrasta contra el spend. 
    Regresa un diccionario con cada columna de actividad y su % de integridad. 
    """
    # Frame con solo las columnas de cada medio declarado
    working_frame = dataframe.filter(like=filter_like)

    # Array para la variable spend
    spend_array = working_frame.filter(like='Spend').notna().to_numpy()
    array_lenght = spend_array.shape[0]

    # Frame para las variables de actividad (Effort)
    working_frame_effort = working_frame.filter(like='Effort')

    # Listado de columnas de actividad (Effort)
    working_frame_effort_columns = working_frame_effort.columns

    # Diccionario para almacenar resultados
    results = {}

    # Iterar por cada columna 
    for effort_column in working_frame_effort_columns:

        # Inicializar la suma 
        sum = 0

        # Obtener el array de la columna de actividad a comparar
        effort_array = working_frame_effort.filter(like=effort_column).notna().to_numpy()
        
        # Hacer la comparacion elemento a elemento 
        for i in range(array_lenght):
            if spend_array[i] == effort_array[i]:
                sum += 1

        # Añadir al diccionario de resultados el ratio de coincidencias / total elementos, es decir % de integridad
        results[effort_column] = sum/array_lenght

    return results


def media_integrity_checker(dataframe, hierarchy_df):
    """
    Funcion que itera por cada medio para identificar el grado de integridad de cada columna de actividad vs su respectivo spend.
    Esta funcion identifica los medios a evaluar y usa la funcion check_integrity para obtener los resultados de integridad de cada columna de actividad. 
    Regresa un dataframe con la integridad por medio. 
    """

    # Crear el listado de medios a evaluar 
    media_list = hierarchy_df[hierarchy_df['L0'] == 'Media'][['L0','L1','L2','L3', 'L4']]
    media_list.drop_duplicates(inplace=True)
    media_list = media_list.agg('::'.join, axis=1).to_list()

    # Diccionario que guarda los resultados
    media_results = {}

    # Itera por cada medio para obtener los resultados de todas las columnas de actividad 
    for media in media_list:
        media_results[media] = check_integrity(dataframe, media)

    return pd.DataFrame(media_results)


def separator_replacer(input_string):
    space_pattern = r'::'
    space_remover = r'\s'
    truncation_pattern = r'(^[^:]*::[^:]*::|::[^:]*::[^:]*::[^:]*$)'

    str_noSpaces = re.sub(pattern=space_remover, repl='', string=input_string)
    str_noSpaces = re.sub(pattern=truncation_pattern, repl='', string=str_noSpaces)

    return re.sub(pattern=space_pattern, repl='_', string=str_noSpaces)


def output_media_column_namer(columns_list):
    """
    Funcion que renombra columnas las variables por cada medio
    """

    trunc_pattern = r'(^[\w\s]*::[\w\s]*::|::[\w\s]*$)'
    space_pattern = r'::'

    output_list = []

    for column_name in columns_list:
        truncated_column_name = re.sub(pattern=trunc_pattern, repl='', string=column_name)
        final_column_name = re.sub(pattern=space_pattern, repl=' ', string=truncated_column_name)

        output_list.append(final_column_name)
    
    return output_list


def per_media_df_generator(dataframe, hierarchy_df):
    """
    Genera un dataframe con la data de cada medio. 
    Regresa tuplas del nombre de la hoja y el dataframe correspondiente 
    """

    # Crear el listado de medios a evaluar 
    media_list = hierarchy_df[hierarchy_df['L0'].apply(lambda x: 'media' in x.lower())][['L0','L1','L2','L3','L4']]
    media_list.drop_duplicates(inplace=True)
    media_list = media_list.agg('::'.join, axis=1).to_list()

    # Lista de dataframes
    media_dataframes = []

    for media in media_list:
        temp_df = dataframe.filter(like=media)
        temp_df.columns = output_media_column_namer(temp_df.columns)
        temp_df.loc[:,'Period'] = dataframe.loc[:,'Period']
        temp_df = temp_df.set_index('Period')

        media_dataframes.append((separator_replacer(media), temp_df))

    return media_dataframes


def cost_per_activity_creator(list_of_per_media_frames):
    """
    Funcion que toma el output de la funcion per_media_df_generator un listado de tuplas, en cada tupla esta el nombre y el dataframe con las variables por medio.
    Regresa un DataFrame con el costo por actividad de cada medio. 
    """

    # Iteramos por cada indice del listado de tuplas 
    flag_first = 0
    for i_media in range(len(list_of_per_media_frames)):

        # Identificamos el nombre de la columna donde esta el spend
        spend_col_name = [name for name in list_of_per_media_frames[i_media][1].columns if 'Spend' in name]

        if len(spend_col_name) == 0:
            continue
        else: 
            flag_first += 1 
        # Si es el primero creamos el dataframe result con la división de todo el frame por la columna del spend 
        if flag_first == 1:
            result_df = list_of_per_media_frames[i_media][1]
            result_df = result_df.apply(lambda col: col / result_df.loc[:,spend_col_name[0]], axis=0)
        
        # Si es una iteracion subsecuente hacemos lo mismo pero vamos haciendo append con el dataframe resultado 
        else:
            temp_df = list_of_per_media_frames[i_media][1]
            temp_df = temp_df.apply(lambda col: col / temp_df.loc[:,spend_col_name[0]], axis=0)
            result_df = pd.merge(result_df, temp_df, left_index=True, right_index=True)
    
    # Hacemos limpieza identificando las columnas de spend y eliminandolas del frame final 
    spend_cols = [name for name in result_df.columns if 'Spend' in name]
    result_df.drop(columns=spend_cols, inplace=True)

    return result_df


def topline_sales_revenue_generator(dataframe_stacked):

    # Filtramos por los datos de ventas en revenue
    filtered_dataframe_sales_rev = dataframe_stacked[(dataframe_stacked['L0'] == 'Sales') & (dataframe_stacked['Metric'] == '$')]

    # Creamos el crosstab de las ventas en revenue por periodo 
    sales_revenue_topline = pd.crosstab(
        index=filtered_dataframe_sales_rev['L1'], 
        columns=filtered_dataframe_sales_rev['Breaks'], 
        values=filtered_dataframe_sales_rev['value'], 
        aggfunc='sum').round(0)

    # Creamos el crosstab del cambio porcentual periodo a periodo 
    sales_revenue_topline_changeYoY = (sales_revenue_topline.T.diff(1) / sales_revenue_topline.T.shift(1)).T.round(4)

    # Lo unimos en una sola tabla
    sales_revenue_dataframe = pd.merge(
        left=sales_revenue_topline, 
        right=sales_revenue_topline_changeYoY,
        left_on='L1',
        right_on='L1',
        suffixes=('',' change'))
    
    return sales_revenue_dataframe


def topline_sales_volume_generator(dataframe_stacked):
    """
    Funcion que genera tablas de volumen de ventas por cada metrica de volumen que tenemos. 
    Genera una lista con tuplas del nombre de pestaña y el dataframe correspondiente. 
    """

    # Filtramos el dataframe principal para tener solo datos de ventas y las metricas que no es revenue. 
    filtered_dataframe_sales_vol = dataframe_stacked[(dataframe_stacked['L0'] == 'Sales') & (dataframe_stacked['Metric'] != '$')]

    # Creamos la lista para capturar la salida de la función 
    sales_vol_dataframe_list = []

    # Por cada metrica ... 
    for metric in filtered_dataframe_sales_vol['Metric'].unique():

        # Volvemos a filtrar el frame por esa submetrica 
        filtered_dataframe_vol_submetric = filtered_dataframe_sales_vol[filtered_dataframe_sales_vol['Metric']==metric]

        # Creamos el crosstab del volumen por periodo (break)
        sales_metric_topline = pd.crosstab(
            index=filtered_dataframe_vol_submetric['L1'], 
            columns=filtered_dataframe_vol_submetric['Breaks'], 
            values=filtered_dataframe_vol_submetric['value'], 
            aggfunc='sum').round(0)

        # Creamos el crosstab del cambio porcentual por periodo 
        sales_metric_topline_changeYoY = (sales_metric_topline.T.diff(1) / sales_metric_topline.T.shift(1)).T.round(4)

        # Los unimos en usa sola tabla
        topline_sales_vol = pd.merge(
            left=sales_metric_topline, 
            right=sales_metric_topline_changeYoY,
            left_on='L1',
            right_on='L1',
            suffixes=(f' ({metric})',f' change ({metric})'))
        
        # Lo añadimos a nuestra lista 
        sales_vol_dataframe_list.append(
            (f'TL Sales Vol {metric}', topline_sales_vol)
        )
    
    # Regresa la salida 
    return sales_vol_dataframe_list 


def topline_efforts_avg_generator(dataframe_stacked):    
    other_efforts_data = dataframe_stacked[(dataframe_stacked['L0'] != 'Media') & (dataframe_stacked['Aggregation'] == 'Average')]

    sheet_list = []

    for effort in other_efforts_data['L0'].unique():

        other_efforts_data_subeffort = other_efforts_data[other_efforts_data['L0'] == effort]

        topline_avg_efforts_cross = pd.crosstab(
            index=other_efforts_data_subeffort['L1'],
            columns=other_efforts_data_subeffort['Breaks'],
            values=other_efforts_data_subeffort['value'],
            aggfunc='mean'
        )

        topline_avg_efforts_cross_changeYoY = (topline_avg_efforts_cross.T.diff(1) / topline_avg_efforts_cross.T.shift(1)).T.round(4)

        topline_avg_efforts_cross = pd.merge(
            left=topline_avg_efforts_cross,
            right=topline_avg_efforts_cross_changeYoY,
            left_on='L1',
            right_on='L1',
            suffixes=('', ' change')
        )

        sheet_list.append((f'{effort}', topline_avg_efforts_cross))

    return sheet_list


def topline_spend_generator(dataframe_stacked):

    # Filtramos por los datos de ventas en revenue
    filtered_dataframe_media = dataframe_stacked[(dataframe_stacked['L0'] == 'Media') & (dataframe_stacked['Metric Type'] == 'Spend')]

    # Creamos una columna con cada medio
    filtered_dataframe_media['L2:L4'] = filtered_dataframe_media.loc[:,('L2','L3','L4')].agg('-'.join, axis=1)

    # Creamos el crosstab de las ventas en revenue por periodo 
    media_spend_dataframe = pd.crosstab(
        index=filtered_dataframe_media['L2:L4'],
        columns=filtered_dataframe_media['Breaks'],
        values=filtered_dataframe_media['value'],
        aggfunc='sum').round(0)

    # Creamos el crosstab del cambio porcentual periodo a periodo 
    media_spend_dataframe_changeYoY = (media_spend_dataframe.T.diff(1) / media_spend_dataframe.T.shift(1)).T.round(4)

    # Creamos el crosstab del market share
    media_spend_dataframe_share = (media_spend_dataframe / media_spend_dataframe.sum()).round(4)

    # Lo unimos en una sola tabla
    media_spend_dataframe = pd.merge(
        left=media_spend_dataframe, 
        right=media_spend_dataframe_changeYoY,
        left_on='L2:L4',
        right_on='L2:L4',
        suffixes=('',' change'))
    
    media_spend_dataframe = pd.merge(
        left=media_spend_dataframe, 
        right=media_spend_dataframe_share,
        left_on='L2:L4',
        right_on='L2:L4',
        suffixes=('',' Spend Share'))
    
    return media_spend_dataframe


def topline_activity_generator_sum(dataframe_stacked):

    # Filtramos por los datos de ventas en revenue
    filtered_dataframe_media = dataframe_stacked[(dataframe_stacked['L0'].apply(lambda x: 'media' in x.lower())) & (dataframe_stacked['Metric Type'] == 'Effort') & (dataframe_stacked['Aggregation'] == 'Sum')]

    # Creamos una columna con cada medio
    filtered_dataframe_media['L0:Metric'] = filtered_dataframe_media.loc[:,('L0','L1','L2','L3','L4','Metric')].agg('-'.join, axis=1)

    # Creamos el crosstab de las ventas en revenue por periodo 
    media_activity_dataframe = pd.crosstab(
        index=filtered_dataframe_media['L0:Metric'],
        columns=filtered_dataframe_media['Breaks'],
        values=filtered_dataframe_media['value'],
        aggfunc='sum').round(0)

    # Creamos el crosstab del cambio porcentual periodo a periodo 
    media_activity_dataframe_changeYoY = (media_activity_dataframe.T.diff(1) / media_activity_dataframe.T.shift(1)).T.round(4)

    # Lo unimos en una sola tabla
    media_activity_dataframe = pd.merge(
        left=media_activity_dataframe, 
        right=media_activity_dataframe_changeYoY,
        left_on='L0:Metric',
        right_on='L0:Metric',
        suffixes=('',' change'))
    
    return media_activity_dataframe


def topline_activity_generator_avg(dataframe_stacked):

    # Filtramos por los datos de ventas en revenue
    filtered_dataframe_media = dataframe_stacked[(dataframe_stacked['L0'].apply(lambda x: 'media' in x.lower())) & (dataframe_stacked['Metric Type'] == 'Effort') & (dataframe_stacked['Aggregation'] == 'Average')]

    # Checamos si el frame esta vacio, si es el caso regresa un frame vacío para que no arroje error la función:
    if filtered_dataframe_media.shape[0]==0:
        return pd.DataFrame()

    # Creamos una columna con cada medio
    filtered_dataframe_media['L0:Metric'] = filtered_dataframe_media.loc[:,('L0','L1','L2','L3','L4','Metric')].agg('-'.join, axis=1)

    # Creamos el crosstab de las ventas en revenue por periodo 
    media_activity_dataframe = pd.crosstab(
        index=filtered_dataframe_media['L0:Metric'],
        columns=filtered_dataframe_media['Breaks'],
        values=filtered_dataframe_media['value'],
        aggfunc='mean').round(0)

    # Creamos el crosstab del cambio porcentual periodo a periodo 
    media_activity_dataframe_changeYoY = (media_activity_dataframe.T.diff(1) / media_activity_dataframe.T.shift(1)).T.round(4)

    # Lo unimos en una sola tabla
    media_activity_dataframe = pd.merge(
        left=media_activity_dataframe, 
        right=media_activity_dataframe_changeYoY,
        left_on='L0:Metric',
        right_on='L0:Metric',
        suffixes=('',' change'))
    
    return media_activity_dataframe


def check_non_numeric_columns(df, exclude_cols):
    """
    Drops specified columns from the DataFrame and returns a list of columns
    that are not of type float32, float64, int32, or int64, along with their types.

    Parameters:
    - df: pandas DataFrame
    - exclude_cols: list of column names to drop

    Returns:
    - List of tuples: (column_name, dtype) for non-numeric columns
    """
    # Drop the specified columns
    df_filtered = df.drop(columns=exclude_cols, errors='ignore')

    # Define allowed numeric types
    allowed_types = ['float32', 'float64', 'int32', 'int64']

    # Identify non-numeric columns
    non_numeric = [(col, str(dtype)) for col, dtype in df_filtered.dtypes.items() if str(dtype) not in allowed_types]

    if len(non_numeric) == 0:
        print('Data OK')
    else:
        print('WARNING: The following columns are non-numeric type, check input!')
        return non_numeric
    

def check_hierarchy_inconsistency(stacked_df, check_col):
    """
    Valida que todas las columnas hayan sido reconocidas en la jerarquia
    """
    na_counts = stacked_df[check_col].isna().sum().sum()
    
    if na_counts > 0:
        print('Check the following columns, there is incosistency between column names in data and heirarchy')
        return stacked_df[stacked_df[check_col].isna()]['variable'].unique()
    
    else:
        print('Stacked data OK')


def sheet_index_creator(list_of_sheet_names=[]):
    index_df = pd.DataFrame(list_of_sheet_names)
    index_df.columns = ['Sheet']
    index_df['Link'] = index_df['Sheet'].apply(lambda sheet_name: f'=HYPERLINK(\"#\'{sheet_name}\'!A1\", \"Go to\")')
    index_df['Notes'] = ''
    
    return index_df