#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Clasificación ABC para el Sistema de Pedidos - Vivero Aranjuez
=========================================================================

Este script realiza un análisis ABC de las compras históricas para identificar
los productos más relevantes según su valor de consumo.

Este script ha sido optimizado para ser completamente portable. Busca los
archivos de entrada en la carpeta 'data/input' ubicada en el mismo directorio
donde se encuentra el script, sin importar en qué ubicación del sistema se
instale.

Author: MiniMax Agent
Date: 2026-02-05
"""

import pandas as pd
import numpy as np
import glob
import os
import sys
from datetime import datetime


def obtener_directorio_base():
    """
    Obtiene el directorio base donde se encuentra el script.
    
    Esta función es la clave de la portabilidad del script. Utiliza
    os.path.abspath(__file__) para obtener la ruta absoluta del archivo
    del script, y luego os.path.dirname() para obtener solo el directorio.
    De esta manera, el script siempre encontrará sus archivos de entrada
    y salida correctamente, sin importar en qué carpeta o unidad se instale.
    
    Returns:
        str: Ruta absoluta al directorio donde se encuentra el script.
    """
    # __file__ contiene la ruta del archivo actual
    # os.path.abspath() convierte cualquier ruta relativa a absoluta
    # os.path.dirname() obtiene el directorio padre del archivo
    return os.path.dirname(os.path.abspath(__file__))


def construir_ruta(relativa):
    """
    Construye una ruta absoluta a partir de una ruta relativa.
    
    Args:
        relativa (str): Ruta relativa al directorio base del script.
    
    Returns:
        str: Ruta absoluta completa.
    """
    dir_base = obtener_directorio_base()
    return os.path.join(dir_base, relativa)


def cargar_datos_compras():
    """
    Carga los datos de compras desde la carpeta de entrada.
    
    Busca el archivo 'compras.xlsx' en la carpeta 'data/input' y lo carga
    en un DataFrame de pandas. Maneja posibles errores de archivo no
    encontrado o formato incorrecto.
    
    Returns:
        pd.DataFrame: DataFrame con los datos de compras, o None si hay error.
    """
    # Construir la ruta al archivo de entrada
    ruta_archivo = construir_ruta("data/input/compras.xlsx")
    
    print(f"Buscando archivo de compras en: {ruta_archivo}")
    
    try:
        # Verificar si el archivo existe
        if not os.path.exists(ruta_archivo):
            print(f"ERROR: No se encontró el archivo '{ruta_archivo}'")
            print("Por favor, verifica que el archivo 'compras.xlsx' existe en la carpeta 'data/input'")
            return None
        
        # Cargar el archivo Excel
        df = pd.read_excel(ruta_archivo)
        print(f"✓ Archivo cargado exitosamente. Total de registros: {len(df)}")
        return df
        
    except Exception as e:
        print(f"ERROR al cargar el archivo: {str(e)}")
        return None


def obtener_archivos_clasificacion():
    """
    Obtiene los archivos existentes de clasificación ABC.
    
    Busca archivos con el patrón 'CLASIFICACION_ABC+D_*.xlsx' en la carpeta
    'data/output' para evitar duplicados o para continuar con análisis
    anteriores.
    
    Returns:
        list: Lista de rutas a archivos encontrados.
    """
    ruta_patron = construir_ruta("data/input/CLASIFICACION_ABC+D_*.xlsx")
    archivos = glob.glob(ruta_patron)
    return archivos


def realizar_clasificacion_abc(df):
    """
    Realiza el análisis ABC sobre los datos de compras.
    
    Calcula el valor acumulado de compras por producto, determina el
    porcentaje acumulado y asigna la categoría ABC según los umbrales
    típicos de distribución (A: 80%, B: 95%, C: resto).
    
    Args:
        df (pd.DataFrame): DataFrame con los datos de compras.
    
    Returns:
        pd.DataFrame: DataFrame con la clasificación ABC añadida.
    """
    print("\nIniciando análisis ABC...")
    
    # Agrupar por producto y sumar el valor de compras
    # Asumiendo que hay columnas 'producto' y 'valor' o similares
    # Si las columnas tienen nombres diferentes, se deben ajustar
    
    # Detectar automáticamente las columnas necesarias
    columnas_df = df.columns.tolist()
    
    # Buscar columna de producto (nombre contiene 'producto', 'articulo', 'item', etc.)
    col_producto = None
    palabras_producto = ['producto', 'articulo', 'item', 'ref', 'referencia', 'código', 'codigo']
    for palabra in palabras_producto:
        for col in columnas_df:
            if palabra in col.lower():
                col_producto = col
                break
        if col_producto:
            break
    
    # Buscar columna de valor (nombre contiene 'valor', 'importe', 'precio', 'cantidad', 'total')
    col_valor = None
    palabras_valor = ['valor', 'importe', 'precio', 'total', 'cantidad']
    for palabra in palabras_valor:
        for col in columnas_df:
            if palabra in col.lower():
                col_valor = col
                break
        if col_valor:
            break
    
    # Si no se encuentran automáticamente, usar las primeras dos columnas
    if col_producto is None:
        col_producto = columnas_df[0] if len(columnas_df) > 0 else None
    if col_valor is None:
        col_valor = columnas_df[1] if len(columnas_df) > 1 else columnas_df[0]
    
    print(f"  - Columna de producto identificada: '{col_producto}'")
    print(f"  - Columna de valor identificada: '{col_valor}'")
    
    # Agrupar y calcular totales por producto
    df_agrupado = df.groupby(col_producto)[col_valor].sum().reset_index()
    df_agrupado.columns = ['Producto', 'Valor_Total']
    
    # Calcular valor acumulado y porcentaje
    df_agrupado = df_agrupado.sort_values('Valor_Total', ascending=False)
    df_agrupado['Valor_Acumulado'] = df_agrupado['Valor_Total'].cumsum()
    total_general = df_agrupado['Valor_Total'].sum()
    df_agrupado['Porcentaje_Acumulado'] = (df_agrupado['Valor_Acumulado'] / total_general) * 100
    
    # Asignar categoría ABC según umbrales
    # Categoría A: 80% del valor (pocos productos representan mucho valor)
    # Categoría B: siguiente 15% (productos moderadamente importantes)
    # Categoría C: resto (muchos productos con poco valor individual)
    def asignar_categoria(porcentaje):
        if porcentaje <= 80:
            return 'A'
        elif porcentaje <= 95:
            return 'B'
        else:
            return 'C'
    
    df_agrupado['Categoria'] = df_agrupado['Porcentaje_Acumulado'].apply(asignar_categoria)
    
    # Calcular estadísticas por categoría
    stats_categoria = df_agrupado.groupby('Categoria').agg({
        'Producto': 'count',
        'Valor_Total': 'sum'
    }).reset_index()
    stats_categoria['Porcentaje_Valor'] = (stats_categoria['Valor_Total'] / total_general) * 100
    
    print("\nResultados del análisis ABC:")
    print("=" * 50)
    for _, row in stats_categoria.iterrows():
        print(f"  Categoría {row['Categoria']}: {row['Producto']:3.0f} productos "
              f"({row['Porcentaje_Valor']:5.1f}% del valor total)")
    print("=" * 50)
    
    return df_agrupado


def guardar_resultados(df_resultados, dir_base):
    """
    Guarda los resultados del análisis en un archivo Excel.
    
    Args:
        df_resultados (pd.DataFrame): DataFrame con los resultados del análisis.
        dir_base (str): Directorio base del script.
    
    Returns:
        str: Ruta al archivo guardado, o None si hay error.
    """
    # Construir ruta de salida con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"CLASIFICACION_ABC+D_{timestamp}.xlsx"
    ruta_salida = construir_ruta(f"data/input/{nombre_archivo}")
    
    # Verificar que existe el directorio de salida
    directorio_salida = construir_ruta("data/input")
    if not os.path.exists(directorio_salida):
        os.makedirs(directorio_salida)
        print(f"✓ Directorio de salida creado: {directorio_salida}")
    
    try:
        # Guardar archivo Excel con formato
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            # Hoja principal con la clasificación
            df_resultados.to_excel(writer, sheet_name='Clasificacion_ABC', index=False)
            
            # Hoja con resumen por categoría
            resumen = df_resultados.groupby('Categoria').agg({
                'Producto': 'count',
                'Valor_Total': ['sum', 'mean', 'min', 'max']
            }).round(2)
            resumen.columns = ['Num_Productos', 'Valor_Total', 'Valor_Promedio', 
                              'Valor_Minimo', 'Valor_Maximo']
            resumen['Porcentaje_Productos'] = (resumen['Num_Productos'] / 
                                                 resumen['Num_Productos'].sum() * 100).round(2)
            resumen.to_excel(writer, sheet_name='Resumen_Por_Categoria')
        
        print(f"\n✓ Resultados guardados en: {ruta_salida}")
        return ruta_salida
        
    except Exception as e:
        print(f"ERROR al guardar resultados: {str(e)}")
        return None


def main():
    """
    Función principal del script.
    
    Orquesta el proceso completo de clasificación ABC:
    1. Carga los datos de compras
    2. Realiza el análisis ABC
    3. Guarda los resultados
    
    Returns:
        int: 0 si el proceso es exitoso, 1 si hay errores.
    """
    print("\n" + "=" * 60)
    print("    CLASIFICACIÓN ABC - VIVEROS ARANJUEZ")
    print("    Sistema Portátil de Análisis de Compras")
    print("=" * 60)
    
    # Mostrar información sobre la ubicación del script
    dir_base = obtener_directorio_base()
    print(f"\nDirectorio del script: {dir_base}")
    print(f"Directorio de entrada: {construir_ruta('data/input')}")
    print(f"Directorio de salida: {construir_ruta('data/input')}")
    
    # Cargar datos
    df = cargar_datos_compras()
    if df is None:
        return 1
    
    # Verificar que hay datos suficientes
    if len(df) == 0:
        print("ERROR: El archivo de compras está vacío")
        return 1
    
    # Realizar clasificación ABC
    df_resultados = realizar_clasificacion_abc(df)
    
    # Guardar resultados
    ruta_archivo = guardar_resultados(df_resultados, dir_base)
    
    if ruta_archivo:
        print("\n✓ Proceso completado exitosamente")
        return 0
    else:
        print("\n✗ Error al guardar los resultados")
        return 1


if __name__ == "__main__":
    sys.exit(main())
