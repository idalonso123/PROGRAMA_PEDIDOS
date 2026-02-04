#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PRESENTACION.py
Script para generar presentaciones HTML del análisis ABC+D de Vivero Aranjuez.
Versión CORREGIDA: Lee SOLO los archivos de clasificación ABC+D y genera presentaciones por sección.
NO utiliza los archivos individuales (Ventas.xlsx, stock.xlsx, compras.xlsx).
"""

import pandas as pd
import numpy as np
from datetime import datetime
import glob
import warnings
import os
warnings.filterwarnings('ignore')


# ============================================================================
# CONFIGURACIÓN
# ============================================================================


def obtener_archivos_clasificacion():
    """Busca todos los archivos de clasificación ABC+D por sección en la carpeta actual."""
    patrones = ["CLASIFICACION_ABC+D_*.xlsx"]
    
    archivos_encontrados = []
    for patron in patrones:
        archivos_encontrados.extend(glob.glob(patron))
    
    # Normalizar rutas y eliminar duplicados
    archivos_normalizados = set()
    archivos_unicos = []
    for archivo in archivos_encontrados:
        archivo_norm = os.path.normpath(archivo)
        if archivo_norm not in archivos_normalizados:
            archivos_normalizados.add(archivo_norm)
            archivos_unicos.append(archivo)
    
    return sorted(archivos_unicos)


def extraer_nombre_seccion(nombre_archivo):
    """Extrae el nombre de la sección del nombre del archivo."""
    basename = os.path.basename(nombre_archivo)
    nombre_sin_extension = basename.replace('.xlsx', '')
    prefijo = "CLASIFICACION_ABC+D_"
    
    if nombre_sin_extension.startswith(prefijo):
        nombre_seccion = nombre_sin_extension[len(prefijo):]
        return nombre_seccion
    
    if 'CLASIFICACION_ABC+D_' in nombre_sin_extension:
        partes = nombre_sin_extension.split('CLASIFICACION_ABC+D_', 1)
        if len(partes) > 1:
            return partes[1]
    
    return None


def leer_datos_clasificacion(ruta_archivo):
    """
    Lee todas las hojas de clasificación del archivo Excel y las combina.
    El archivo de clasificación YA contiene los datos calculados correctamente.
    """
    excel_file = pd.ExcelFile(ruta_archivo)
    hojas = {}
    df_combinado = None
    
    for hoja in excel_file.sheet_names:
        df_hoja = pd.read_excel(excel_file, sheet_name=hoja)
        hojas[hoja] = df_hoja
        
        # Combinar todas las hojas
        if df_combinado is None:
            df_combinado = df_hoja
        else:
            df_combinado = pd.concat([df_combinado, df_hoja], ignore_index=True)
    
    return hojas, df_combinado


def obtener_datos_seccion(hojas_dict):
    """
    Obtiene los datos consolidados de la sección desde las hojas de clasificación.
    USA LOS DATOS YA CALCULADOS del archivo de clasificación.
    """
    # Combinar todas las categorías
    df_seccion_completo = pd.concat(hojas_dict.values(), ignore_index=True)
    
    # Calcular métricas usando las columnas pre-calculadas del archivo de clasificación
    # El archivo CLASIFICACION_ABC+D ya tiene las columnas correctas:
    # - 'Importe ventas (€)' o 'Importe Ventas'
    # - 'Beneficio (importe €)' o 'Beneficio'
    # - 'Stock' o 'Stock (uds)'
    
    # Detectar nombres de columnas
    col_articulo = 'Artículo'
    col_ventas = 'Importe ventas (€)'
    col_beneficio = 'Beneficio (importe €)'
    col_stock = 'Stock Final (unidades)'
    
    # Calcular resumen de la sección
    datos_seccion = {
        'total_articulos': int(len(df_seccion_completo)),
        'ventas_totales': float(df_seccion_completo[col_ventas].sum()),
        'beneficio_total': float(df_seccion_completo[col_beneficio].sum()),
        'stock_final_total': int(df_seccion_completo[col_stock].sum())
    }
    
    # Calcular margen bruto (dato ya viene pre-calculado en el archivo de clasificación)
    if datos_seccion['ventas_totales'] > 0:
        datos_seccion['margen_bruto'] = round(datos_seccion['beneficio_total'] / datos_seccion['ventas_totales'] * 100, 1)
    else:
        datos_seccion['margen_bruto'] = 0
    
    # Calcular distribución por categoría
    dist_categoria = df_seccion_completo.groupby(col_articulo).agg({
        col_ventas: 'sum',
        col_beneficio: 'sum',
        col_stock: 'sum'
    }).reset_index()
    
    # Contar artículos por categoría basándose en los nombres de las hojas
    # Usamos un patrón más específico para evitar coincidencias parciales
    categorias = {}
    ventas_por_categoria = {}
    stock_por_categoria = {}
    
    for nombre_hoja, df_hoja in hojas_dict.items():
        nombre_upper = nombre_hoja.upper()
        if 'CATEGORIA A' in nombre_upper or ' A –' in nombre_upper or nombre_upper.startswith('A –'):
            categorias['A'] = len(df_hoja)
            ventas_por_categoria['A'] = float(df_hoja[col_ventas].sum())
            stock_por_categoria['A'] = int(df_hoja[col_stock].sum())
        elif 'CATEGORIA B' in nombre_upper or ' B –' in nombre_upper or nombre_upper.startswith('B –'):
            categorias['B'] = len(df_hoja)
            ventas_por_categoria['B'] = float(df_hoja[col_ventas].sum())
            stock_por_categoria['B'] = int(df_hoja[col_stock].sum())
        elif 'CATEGORIA C' in nombre_upper or ' C –' in nombre_upper or nombre_upper.startswith('C –'):
            categorias['C'] = len(df_hoja)
            ventas_por_categoria['C'] = float(df_hoja[col_ventas].sum())
            stock_por_categoria['C'] = int(df_hoja[col_stock].sum())
        elif 'CATEGORIA D' in nombre_upper or ' D –' in nombre_upper or nombre_upper.startswith('D –'):
            categorias['D'] = len(df_hoja)
            ventas_por_categoria['D'] = float(df_hoja[col_ventas].sum())
            stock_por_categoria['D'] = int(df_hoja[col_stock].sum())
    
    return datos_seccion, categorias, ventas_por_categoria, stock_por_categoria


def formatear_numero(valor, decimales=0):
    """Formatea un número con separadores de miles."""
    if valor is None:
        return "0"
    try:
        return f"{float(valor):,.{decimales}f}"
    except:
        return "0"


def generar_html_presentacion(datos_seccion, categorias, ventas_por_categoria, stock_por_categoria, nombre_seccion=None):
    """
    Genera el HTML de la presentación interactiva.
    """
    fecha_actual = datetime.now().strftime("%d de %B de %Y")
    nombre_seccion_titulo = nombre_seccion.upper() if nombre_seccion else "VIVERO ARANJUEZ"
    
    # Obtener valores por categoría
    count_a = categorias.get('A', 0)
    count_b = categorias.get('B', 0)
    count_c = categorias.get('C', 0)
    count_d = categorias.get('D', 0)
    
    ventas_a = ventas_por_categoria.get('A', 0)
    ventas_b = ventas_por_categoria.get('B', 0)
    ventas_c = ventas_por_categoria.get('C', 0)
    ventas_d = ventas_por_categoria.get('D', 0)
    
    stock_a = stock_por_categoria.get('A', 0)
    stock_b = stock_por_categoria.get('B', 0)
    stock_c = stock_por_categoria.get('C', 0)
    stock_d = stock_por_categoria.get('D', 0)
    
    total_arts = count_a + count_b + count_c + count_d
    total_ventas = ventas_a + ventas_b + ventas_c + ventas_d
    
    # Calcular porcentajes
    pct_a = round(count_a / total_arts * 100, 1) if total_arts > 0 else 0
    pct_b = round(count_b / total_arts * 100, 1) if total_arts > 0 else 0
    pct_c = round(count_c / total_arts * 100, 1) if total_arts > 0 else 0
    pct_d = round(count_d / total_arts * 100, 1) if total_arts > 0 else 0
    
    pct_ventas_a = round(ventas_a / total_ventas * 100, 1) if total_ventas > 0 else 0
    pct_ventas_b = round(ventas_b / total_ventas * 100, 1) if total_ventas > 0 else 0
    pct_ventas_c = round(ventas_c / total_ventas * 100, 1) if total_ventas > 0 else 0
    
    # Calcular ángulos para el gráfico (distribución por artículos)
    angle_a = round(pct_a * 3.6, 1)
    angle_b = round(pct_b * 3.6, 1)
    angle_c = round(pct_c * 3.6, 1)
    angle_d = round(360 - angle_a - angle_b - angle_c, 1)
    
    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Presentación - Sección {nombre_seccion_titulo} | ABC+D</title>
    <style>
        :root {{
            --primary: #2d5a27;
            --primary-light: #4a8c3f;
            --accent: #8bc34a;
            --warning: #ff9800;
            --danger: #f44336;
            --success: #4caf50;
            --info: #2196f3;
            --dark: #1a1a1a;
            --light: #f8f9fa;
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: var(--dark);
            color: var(--dark);
            overflow: hidden;
            height: 100vh;
        }}
        
        .presentation {{
            width: 100%;
            height: 100vh;
            position: relative;
        }}
        
        .slide {{
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            display: flex;
            flex-direction: column;
            opacity: 0;
            visibility: hidden;
            transition: all 0.6s cubic-bezier(0.4, 0, 0.2, 1);
            transform: translateX(100px);
            background: linear-gradient(135deg, #ffffff 0%, #f5f9f5 100%);
            padding: 60px 80px;
        }}
        
        .slide.active {{
            opacity: 1;
            visibility: visible;
            transform: translateX(0);
        }}
        
        .slide.prev {{
            transform: translateX(-100px);
        }}
        
        .slide-title {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
            color: white;
            justify-content: center;
            align-items: center;
            text-align: center;
        }}
        
        .slide-title h1 {{
            font-size: 3.5em;
            font-weight: 700;
            margin-bottom: 20px;
            text-shadow: 2px 2px 10px rgba(0,0,0,0.2);
        }}
        
        .slide-title .subtitle {{
            font-size: 1.6em;
            opacity: 0.9;
            margin-bottom: 30px;
        }}
        
        .slide-title .meta {{
            font-size: 1.1em;
            opacity: 0.8;
        }}
        
        .slide-title .section-badge {{
            background: rgba(255,255,255,0.2);
            padding: 10px 30px;
            border-radius: 30px;
            font-size: 1.3em;
            margin-bottom: 30px;
        }}
        
        .slide-header {{
            display: flex;
            align-items: center;
            gap: 20px;
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 3px solid var(--accent);
        }}
        
        .slide-header .icon {{
            width: 60px;
            height: 60px;
            background: var(--primary);
            color: white;
            border-radius: 15px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.8em;
        }}
        
        .slide-header h2 {{
            font-size: 2em;
            color: var(--primary);
        }}
        
        .slide-content {{
            flex: 1;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        
        .kpi-row {{
            display: flex;
            gap: 25px;
            justify-content: center;
            flex-wrap: wrap;
        }}
        
        .kpi-box {{
            background: white;
            padding: 30px 40px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            min-width: 200px;
            transition: transform 0.3s ease;
        }}
        
        .kpi-box:hover {{
            transform: translateY(-8px);
        }}
        
        .kpi-box .number {{
            font-size: 3em;
            font-weight: 800;
            color: var(--primary);
            line-height: 1;
        }}
        
        .kpi-box .label {{
            font-size: 1em;
            color: #666;
            margin-top: 8px;
        }}
        
        .kpi-box.highlight {{
            background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
        }}
        
        .kpi-box.highlight .number,
        .kpi-box.highlight .label {{
            color: white;
        }}
        
        .category-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            height: 100%;
        }}
        
        .cat-card {{
            border-radius: 15px;
            padding: 25px;
            color: white;
            display: flex;
            flex-direction: column;
            justify-content: center;
            position: relative;
            overflow: hidden;
            transition: transform 0.3s ease;
        }}
        
        .cat-card:hover {{
            transform: scale(1.02);
        }}
        
        .cat-card::before {{
            content: '';
            position: absolute;
            top: -50%;
            right: -50%;
            width: 100%;
            height: 100%;
            background: rgba(255,255,255,0.1);
            border-radius: 50%;
        }}
        
        .cat-card.a {{
            background: linear-gradient(135deg, #1b5e20 0%, #4caf50 100%);
        }}
        
        .cat-card.b {{
            background: linear-gradient(135deg, #1565c0 0%, #42a5f5 100%);
        }}
        
        .cat-card.c {{
            background: linear-gradient(135deg, #e65100 0%, #ff9800 100%);
        }}
        
        .cat-card.d {{
            background: linear-gradient(135deg, #c62828 0%, #f44336 100%);
        }}
        
        .cat-card h3 {{
            font-size: 1.3em;
            margin-bottom: 10px;
        }}
        
        .cat-card .count {{
            font-size: 2.5em;
            font-weight: 800;
            margin: 10px 0;
        }}
        
        .cat-card .details {{
            font-size: 0.9em;
            opacity: 0.9;
            line-height: 1.5;
        }}
        
        .cat-card .badge {{
            position: absolute;
            top: 15px;
            right: 15px;
            background: rgba(255,255,255,0.2);
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.8em;
            font-weight: 600;
        }}
        
        .chart-section {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 60px;
            flex: 1;
        }}
        
        .donut-chart {{
            width: 300px;
            height: 300px;
            border-radius: 50%;
            background: conic-gradient(
                #1b5e20 0deg {angle_a}deg,
                #1565c0 {angle_a}deg {angle_a + angle_b}deg,
                #e65100 {angle_a + angle_b}deg {angle_a + angle_b + angle_c}deg,
                #c62828 {angle_a + angle_b + angle_c}deg 360deg
            );
            position: relative;
            box-shadow: 0 20px 60px rgba(0,0,0,0.2);
        }}
        
        .donut-chart::before {{
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 150px;
            height: 150px;
            background: linear-gradient(135deg, #ffffff 0%, #f5f9f5 100%);
            border-radius: 50%;
        }}
        
        .chart-center {{
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            z-index: 1;
        }}
        
        .chart-center .total {{
            font-size: 2em;
            font-weight: 800;
            color: var(--primary);
        }}
        
        .chart-center .label {{
            font-size: 0.9em;
            color: #666;
        }}
        
        .chart-legend {{
            display: flex;
            flex-direction: column;
            gap: 15px;
        }}
        
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 15px;
            padding: 12px 20px;
            background: white;
            border-radius: 10px;
            box-shadow: 0 3px 15px rgba(0,0,0,0.08);
        }}
        
        .legend-color {{
            width: 20px;
            height: 20px;
            border-radius: 5px;
        }}
        
        .legend-text h4 {{
            font-size: 1.1em;
            color: var(--dark);
        }}
        
        .legend-text p {{
            font-size: 0.85em;
            color: #666;
        }}
        
        .key-points {{
            display: flex;
            flex-direction: column;
            gap: 18px;
        }}
        
        .key-point {{
            display: flex;
            align-items: center;
            gap: 20px;
            background: white;
            padding: 20px 25px;
            border-radius: 12px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.08);
        }}
        
        .key-point .icon-box {{
            width: 50px;
            height: 50px;
            background: var(--bg-light);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5em;
            flex-shrink: 0;
        }}
        
        .key-point h4 {{
            font-size: 1.1em;
            color: var(--dark);
            margin-bottom: 3px;
        }}
        
        .key-point p {{
            color: #666;
            font-size: 0.9em;
        }}
        
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 25px;
        }}
        
        .summary-box {{
            background: white;
            border-radius: 15px;
            padding: 25px;
            text-align: center;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }}
        
        .summary-box h3 {{
            font-size: 1.1em;
            color: var(--primary);
            margin-bottom: 15px;
        }}
        
        .summary-box .big-number {{
            font-size: 2.5em;
            font-weight: 800;
            color: var(--primary);
        }}
        
        .summary-box p {{
            color: #666;
            margin-top: 8px;
            font-size: 0.9em;
        }}
        
        .nav-controls {{
            position: fixed;
            bottom: 30px;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 15px;
            z-index: 1000;
        }}
        
        .nav-btn {{
            width: 55px;
            height: 55px;
            border: none;
            background: var(--primary);
            color: white;
            border-radius: 50%;
            cursor: pointer;
            font-size: 1.4em;
            transition: all 0.3s ease;
            box-shadow: 0 5px 20px rgba(0,0,0,0.2);
        }}
        
        .nav-btn:hover {{
            background: var(--primary-light);
            transform: scale(1.1);
        }}
        
        .nav-btn:disabled {{
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }}
        
        .slide-counter {{
            position: fixed;
            bottom: 30px;
            right: 30px;
            background: var(--dark);
            color: white;
            padding: 8px 18px;
            border-radius: 25px;
            font-weight: 600;
            z-index: 1000;
        }}
        
        .progress-bar {{
            position: fixed;
            top: 0;
            left: 0;
            height: 5px;
            background: var(--accent);
            transition: width 0.3s ease;
            z-index: 1000;
        }}
        
        @media (max-width: 1200px) {{
            .category-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
            .chart-section {{
                flex-direction: column;
            }}
        }}
        
        @media (max-width: 768px) {{
            .slide {{
                padding: 30px 40px;
            }}
            .category-grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <div class="progress-bar" id="progressBar"></div>
    
    <div class="presentation">
        <!-- Slide 1: Title -->
        <div class="slide slide-title active" data-slide="1">
            <div class="section-badge">{nombre_seccion_titulo}</div>
            <h1>Clasificación ABC+D</h1>
            <p class="subtitle">Análisis de Inventario y Ventas</p>
            <p class="meta">Vivero Aranjuez | {fecha_actual}</p>
        </div>
        
        <!-- Slide 2: Agenda -->
        <div class="slide" data-slide="2">
            <div class="slide-header">
                <div class="icon">📋</div>
                <h2>Agenda de Presentación</h2>
            </div>
            <div class="slide-content">
                <div class="key-points">
                    <div class="key-point">
                        <div class="icon-box">📊</div>
                        <div>
                            <h4>Resultados Generales</h4>
                            <p>Resumen de la clasificación ABC+D y métricas clave</p>
                        </div>
                    </div>
                    <div class="key-point">
                        <div class="icon-box">📦</div>
                        <div>
                            <h4>Distribución por Categorías</h4>
                            <p>Análisis detallado de Categorías A, B, C y D</p>
                        </div>
                    </div>
                    <div class="key-point">
                        <div class="icon-box">💰</div>
                        <div>
                            <h4>Participación en Ingresos</h4>
                            <p>Gráfico de distribución de ventas por categoría</p>
                        </div>
                    </div>
                    <div class="key-point">
                        <div class="icon-box">📈</div>
                        <div>
                            <h4>Acciones Recomendadas</h4>
                            <p>Estrategias para optimizar el inventario</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 3: Key Metrics -->
        <div class="slide" data-slide="3">
            <div class="slide-header">
                <div class="icon">📈</div>
                <h2>Métricas Clave</h2>
            </div>
            <div class="slide-content">
                <div class="kpi-row">
                    <div class="kpi-box highlight">
                        <div class="number">{total_arts}</div>
                        <div class="label">Total Artículos</div>
                    </div>
                    <div class="kpi-box">
                        <div class="number">{formatear_numero(datos_seccion['ventas_totales'], 0)}€</div>
                        <div class="label">Ventas Totales</div>
                    </div>
                    <div class="kpi-box">
                        <div class="number">{datos_seccion['stock_final_total']:,}</div>
                        <div class="label">Unidades Stock</div>
                    </div>
                    <div class="kpi-box">
                        <div class="number">{datos_seccion['margen_bruto']}%</div>
                        <div class="label">Margen Bruto</div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 4: Categories Overview -->
        <div class="slide" data-slide="4">
            <div class="slide-header">
                <div class="icon">📦</div>
                <h2>Distribución por Categorías</h2>
            </div>
            <div class="slide-content">
                <div class="category-grid">
                    <div class="cat-card a">
                        <span class="badge">{pct_a}%</span>
                        <h3>Categoría A</h3>
                        <div class="count">{count_a}</div>
                        <div class="details">
                            <p>Artículos ({pct_ventas_a}% ventas)</p>
                            <p>Stock: {stock_a} uds</p>
                        </div>
                    </div>
                    <div class="cat-card b">
                        <span class="badge">{pct_b}%</span>
                        <h3>Categoría B</h3>
                        <div class="count">{count_b}</div>
                        <div class="details">
                            <p>Artículos ({pct_ventas_b}% ventas)</p>
                            <p>Stock: {stock_b} uds</p>
                        </div>
                    </div>
                    <div class="cat-card c">
                        <span class="badge">{pct_c}%</span>
                        <h3>Categoría C</h3>
                        <div class="count">{count_c}</div>
                        <div class="details">
                            <p>Artículos ({pct_ventas_c}% ventas)</p>
                            <p>Stock: {stock_c} uds</p>
                        </div>
                    </div>
                    <div class="cat-card d">
                        <span class="badge">{pct_d}%</span>
                        <h3>Categoría D</h3>
                        <div class="count">{count_d}</div>
                        <div class="details">
                            <p>Sin ventas</p>
                            <p>Stock: {stock_d} uds</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 5: Revenue Distribution -->
        <div class="slide" data-slide="5">
            <div class="slide-header">
                <div class="icon">💰</div>
                <h2>Participación en Ingresos</h2>
            </div>
            <div class="slide-content">
                <div class="chart-section">
                    <div class="donut-chart">
                        <div class="chart-center">
                            <div class="total">{formatear_numero(total_ventas, 0)}€</div>
                            <div class="label">Total Ventas</div>
                        </div>
                    </div>
                    <div class="chart-legend">
                        <div class="legend-item">
                            <div class="legend-color" style="background: linear-gradient(135deg, #1b5e20, #4caf50);"></div>
                            <div class="legend-text">
                                <h4>A - {pct_ventas_a}%</h4>
                                <p>{formatear_numero(ventas_a, 0)}€ - Productos Estrella</p>
                            </div>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: linear-gradient(135deg, #1565c0, #42a5f5);"></div>
                            <div class="legend-text">
                                <h4>B - {pct_ventas_b}%</h4>
                                <p>{formatear_numero(ventas_b, 0)}€ - Complemento</p>
                            </div>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: linear-gradient(135deg, #e65100, #ff9800);"></div>
                            <div class="legend-text">
                                <h4>C - {pct_ventas_c}%</h4>
                                <p>{formatear_numero(ventas_c, 0)}€ - Bajo Movimiento</p>
                            </div>
                        </div>
                        <div class="legend-item">
                            <div class="legend-color" style="background: linear-gradient(135deg, #c62828, #f44336);"></div>
                            <div class="legend-text">
                                <h4>D - 0%</h4>
                                <p>Sin ventas</p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 6: Key Findings -->
        <div class="slide" data-slide="6">
            <div class="slide-header">
                <div class="icon">🔑</div>
                <h2>Hallazgos Principales</h2>
            </div>
            <div class="slide-content">
                <div class="key-points">
                    <div class="key-point">
                        <div class="icon-box">📊</div>
                        <div>
                            <h4>Concentración de Valor</h4>
                            <p>El {pct_a}% de artículos genera el {pct_ventas_a}% de ventas (principio de Pareto)</p>
                        </div>
                    </div>
                    <div class="key-point" style="border-left: 5px solid #f44336;">
                        <div class="icon-box">⚠️</div>
                        <div>
                            <h4>Artículos Sin Ventas</h4>
                            <p>{count_d} artículos ({pct_d}%) sin rotación - requieren acción</p>
                        </div>
                    </div>
                    <div class="key-point">
                        <div class="icon-box">💵</div>
                        <div>
                            <h4>Margen de Beneficio</h4>
                            <p>Margen bruto del {datos_seccion['margen_bruto']}% - rentabilidad saludable</p>
                        </div>
                    </div>
                    <div class="key-point">
                        <div class="icon-box">📦</div>
                        <div>
                            <h4>Stock Total</h4>
                            <p>{datos_seccion['stock_final_total']:,} unidades valoradas en el inventario</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 7: Recommendations -->
        <div class="slide" data-slide="7">
            <div class="slide-header">
                <div class="icon">💡</div>
                <h2>Recomendaciones</h2>
            </div>
            <div class="slide-content">
                <div class="key-points">
                    <div class="key-point" style="border-left: 5px solid #1b5e20;">
                        <div class="icon-box">⭐</div>
                        <div>
                            <h4>Priorizar Categoría A</h4>
                            <p>Asegurar stock óptimo de {count_a} artículos estrella - prioridad máxima</p>
                        </div>
                    </div>
                    <div class="key-point" style="border-left: 5px solid #ff9800;">
                        <div class="icon-box">🏷️</div>
                        <div>
                            <h4>Promocionar Categoría B</h4>
                            <p>Aplicar estrategias de promoción para aumentar rotaciones</p>
                        </div>
                    </div>
                    <div class="key-point" style="border-left: 5px solid #e65100;">
                        <div class="icon-box">📉</div>
                        <div>
                            <h4>Evaluar Categoría C</h4>
                            <p>Revisar continuidad de {count_c} artículos con bajo rendimiento</p>
                        </div>
                    </div>
                    <div class="key-point" style="border-left: 5px solid #c62828;">
                        <div class="icon-box">🛒</div>
                        <div>
                            <h4>Liquidar Categoría D</h4>
                            <p>Tomar decisiones sobre {count_d} artículos sin ventas ({pct_d}% del catálogo)</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 8: Summary -->
        <div class="slide" data-slide="8">
            <div class="slide-header">
                <div class="icon">✅</div>
                <h2>Resumen</h2>
            </div>
            <div class="slide-content">
                <div class="summary-grid">
                    <div class="summary-box">
                        <h3>Artículos Activos</h3>
                        <div class="big-number">{count_a + count_b + count_c}</div>
                        <p>En Categorías A, B y C con potencial de ventas</p>
                    </div>
                    <div class="summary-box">
                        <h3>Sin Rotación</h3>
                        <div class="big-number">{count_d}</div>
                        <p>Artículos sin ventas ({pct_d}% del catálogo)</p>
                    </div>
                    <div class="summary-box">
                        <h3>Margen Bruto</h3>
                        <div class="big-number">{datos_seccion['margen_bruto']}%</div>
                        <p>Rentabilidad del negocio</p>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Slide 9: Thank You -->
        <div class="slide slide-title" data-slide="9">
            <div class="section-badge">{nombre_seccion_titulo}</div>
            <h1>¡Gracias!</h1>
            <p class="subtitle">¿Preguntas o comentarios?</p>
            <p class="meta">Vivero Aranjuez | Análisis ABC+D</p>
        </div>
    </div>
    
    <div class="nav-controls">
        <button class="nav-btn" id="prevBtn" onclick="changeSlide(-1)">←</button>
        <button class="nav-btn" id="nextBtn" onclick="changeSlide(1)">→</button>
    </div>
    
    <div class="slide-counter">
        <span id="currentSlide">1</span> / <span id="totalSlides">9</span>
    </div>
    
    <script>
        let currentSlide = 1;
        const totalSlides = document.querySelectorAll('.slide').length;
        
        document.getElementById('totalSlides').textContent = totalSlides;
        
        function updateSlide() {{
            document.querySelectorAll('.slide').forEach((slide, index) => {{
                slide.classList.remove('active', 'prev');
                if (index + 1 === currentSlide) {{
                    slide.classList.add('active');
                }} else if (index + 1 < currentSlide) {{
                    slide.classList.add('prev');
                }}
            }});
            
            document.getElementById('currentSlide').textContent = currentSlide;
            document.getElementById('prevBtn').disabled = currentSlide === 1;
            document.getElementById('nextBtn').disabled = currentSlide === totalSlides;
            
            const progress = ((currentSlide - 1) / (totalSlides - 1)) * 100;
            document.getElementById('progressBar').style.width = progress + '%';
        }}
        
        function changeSlide(direction) {{
            currentSlide += direction;
            if (currentSlide < 1) currentSlide = 1;
            if (currentSlide > totalSlides) currentSlide = totalSlides;
            updateSlide();
        }}
        
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'ArrowRight' || e.key === ' ') {{
                changeSlide(1);
            }} else if (e.key === 'ArrowLeft') {{
                changeSlide(-1);
            }}
        }});
        
        let touchStartX = 0;
        document.addEventListener('touchstart', function(e) {{
            touchStartX = e.touches[0].clientX;
        }});
        
        document.addEventListener('touchend', function(e) {{
            const touchEndX = e.changedTouches[0].clientX;
            const diff = touchStartX - touchEndX;
            if (Math.abs(diff) > 50) {{
                if (diff > 0) {{
                    changeSlide(1);
                }} else {{
                    changeSlide(-1);
                }}
            }}
        }});
        
        updateSlide();
    </script>
</body>
</html>"""
    
    return html


def main():
    """
    Función principal que ejecuta todo el proceso.
    Lee SOLO los archivos de clasificación ABC+D (ya contienen los datos correctos).
    """
    print("=" * 70)
    print("GENERADOR DE PRESENTACIONES ABC+D POR SECCIÓN")
    print("Vivero Aranjuez")
    print("=" * 70)
    print("\nMODO: Usando archivos de clasificación como fuente de datos")
    print("      (NO se procesan archivos individuales Ventas/stock/compras)")
    
    # Buscar archivos de clasificación
    print("\n[1/3] Buscando archivos de clasificación ABC+D...")
    archivos_clasificacion = obtener_archivos_clasificacion()
    
    if not archivos_clasificacion:
        print("    ⚠ No se encontraron archivos CLASIFICACION_ABC+D_*.xlsx")
        print("    Asegúrate de tener los archivos de clasificación en la carpeta.")
        return
    
    print(f"    ✓ Se encontraron {len(archivos_clasificacion)} archivo(s) de clasificación")
    
    # Procesar cada sección
    print("\n[2/3] Procesando secciones...")
    presentaciones_generadas = 0
    errores = 0
    
    for archivo in archivos_clasificacion:
        nombre_seccion = extraer_nombre_seccion(archivo)
        if not nombre_seccion:
            print(f"    ERROR: No se pudo extraer nombre de sección de {archivo}")
            errores += 1
            continue
        
        print(f"\n    Procesando: {nombre_seccion}")
        print(f"    Archivo: {archivo}")
        
        try:
            # Leer datos de clasificación (TODAS las hojas)
            print("    [1/2] Leyendo clasificación...")
            hojas_dict, df_combinado = leer_datos_clasificacion(archivo)
            print(f"      ✓ Hojas leídas: {list(hojas_dict.keys())}")
            print(f"      ✓ Total artículos: {len(df_combinado)}")
            
            # Obtener datos de la sección
            print("    [2/2] Generando presentación...")
            datos_seccion, categorias, ventas_por_categoria, stock_por_categoria = obtener_datos_seccion(hojas_dict)
            
            # Generar HTML
            html_presentacion = generar_html_presentacion(
                datos_seccion, 
                categorias, 
                ventas_por_categoria, 
                stock_por_categoria, 
                nombre_seccion
            )
            
            # Guardar archivo
            nombre_salida = f"PRESENTACION_{nombre_seccion}.html"
            with open(nombre_salida, 'w', encoding='utf-8') as f:
                f.write(html_presentacion)
            
            print(f"      ✓ GENERADO: {nombre_salida}")
            print(f"      ✓ Artículos: {datos_seccion['total_articulos']}")
            print(f"      ✓ Ventas: {formatear_numero(datos_seccion['ventas_totales'], 0)}€")
            print(f"      ✓ Margen: {datos_seccion['margen_bruto']}%")
            
            presentaciones_generadas += 1
            
        except Exception as e:
            print(f"      ERROR: {str(e)}")
            import traceback
            traceback.print_exc()
            errores += 1
    
    # Resumen final
    print("\n" + "=" * 70)
    print("RESUMEN DE GENERACIÓN DE PRESENTACIONES")
    print("=" * 70)
    print(f"  Archivos de clasificación procesados: {len(archivos_clasificacion)}")
    print(f"  Presentaciones generadas: {presentaciones_generadas}")
    print(f"  Errores: {errores}")
    print("=" * 70)
    
    if presentaciones_generadas > 0:
        print("\n✓ PROCESO COMPLETADO EXITOSAMENTE")
        print(f"Se han generado {presentaciones_generadas} presentación(es):")
        for archivo in archivos_clasificacion:
            nombre_seccion = extraer_nombre_seccion(archivo)
            if nombre_seccion:
                print(f"  - PRESENTACION_{nombre_seccion}.html")
    else:
        print("\n⚠ No se generaron presentaciones. Revisa los errores anteriores.")
    
    print("\n" + "=" * 70)
    print("Para ver las presentaciones, abre los archivos HTML en un navegador.")
    print("Navegación: Usa las flechas del teclado o los botones en pantalla.")
    print("=" * 70)


if __name__ == "__main__":
    main()
