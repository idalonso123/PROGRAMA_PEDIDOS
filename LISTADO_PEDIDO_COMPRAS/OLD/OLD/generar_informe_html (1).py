#!/usr/bin/env python3
"""
Script para generar informes ejecutivos HTML para todas las secciones
Genera informes visuales y profesionales con métricas, gráficos y análisis
USANDO LOS DATOS YA CALCULADOS EN Resumen_Pedidos.xlsx

NUEVA FUNCIONALIDAD: Generación automática de informes para todas las secciones
en una sola ejecución, más un informe consolidado global
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings
warnings.filterwarnings('ignore')

# ============================================
# CONFIGURACIÓN DE SECCIONES
# ============================================

# Definición de todas las secciones con nombres para archivos y títulos
SECTIONS_CONFIG = {
    'interior': {
        'nombre_archivo': 'interior',
        'titulo_seccion': 'PLANTAS DE INTERIOR',
        'descripcion': 'Plantas de interior'
    },
    'utiles_jardin': {
        'nombre_archivo': 'utiles_jardin',
        'titulo_seccion': 'ÚTILES DE JARDÍN',
        'descripcion': 'Útiles y herramientas de jardín'
    },
    'semillas': {
        'nombre_archivo': 'semillas',
        'titulo_seccion': 'SEMILLAS Y BULBOS',
        'descripcion': 'Semillas, bulbos y plantar'
    },
    'deco_interior': {
        'nombre_archivo': 'deco_interior',
        'titulo_seccion': 'DECORACIÓN INTERIOR',
        'descripcion': 'Artículos de decoración interior'
    },
    'maf': {
        'nombre_archivo': 'maf',
        'titulo_seccion': 'PLANTA DE TEMPORADA Y FLORISTERÍA',
        'descripcion': 'Planta de temporada y productos de floristería'
    },
    'vivero': {
        'nombre_archivo': 'vivero',
        'titulo_seccion': 'VIVERO Y PLANTAS EXTERIOR',
        'descripcion': 'Plantas de exterior y viveros'
    },
    'deco_exterior': {
        'nombre_archivo': 'deco_exterior',
        'titulo_seccion': 'DECORACIÓN EXTERIOR',
        'descripción': 'Artículos de decoración para exterior'
    },
    'mascotas_manufacturado': {
        'nombre_archivo': 'mascotas_manufacturado',
        'titulo_seccion': 'MASCOTAS (PRODUCTOS)',
        'descripcion': 'Productos para mascotas manufacturados'
    },
    'mascotas_vivo': {
        'nombre_archivo': 'mascotas_vivo',
        'titulo_seccion': 'ANIMALES VIVOS',
        'descripcion': 'Animales vivos para venta'
    },
    'tierra_aridos': {
        'nombre_archivo': 'tierra_aridos',
        'titulo_seccion': 'TIERRAS Y ÁRIDOS',
        'descripcion': 'Tierras, sustratos y áridos'
    },
    'fitos': {
        'nombre_archivo': 'fitos',
        'titulo_seccion': 'FITOSANITARIOS Y ABONOS',
        'descripcion': 'Fitosanitarios, abonos y fertilizantes'
    }
}

# ============================================
# DETECCIÓN AUTOMÁTICA DE SECCIÓN
# Copiado exactamente de pedido_compras.py
# ============================================

# Códigos de animales vivos (tienen tratamiento especial dentro de sección 2)
CODIGOS_MASCOTAS_VIVO = ['2104', '2204', '2305', '2405', '2504', '2606', '2705', '2707', '2708', '2805', '2806', '2906']

# Definición de todas las secciones y sus rangos de códigos
SECCIONES = {
    'interior': {
        'descripcion': 'Plantas de interior',
        'rangos': [{'tipo': 'prefijos', 'valores': ['1']}]
    },
    'utiles_jardin': {
        'descripcion': 'Útiles de jardín',
        'rangos': [{'tipo': 'prefijos', 'valores': ['4']}]
    },
    'semillas': {
        'descripcion': 'Semillas y bulbos',
        'rangos': [{'tipo': 'prefijos', 'valores': ['5']}]
    },
    'deco_interior': {
        'descripcion': 'Decoración interior',
        'rangos': [{'tipo': 'prefijos', 'valores': ['6']}]
    },
    'maf': {
        'descripcion': 'Planta de temporada y floristería',
        'rangos': [{'tipo': 'prefijos', 'valores': ['7']}]
    },
    'vivero': {
        'descripcion': 'Vivero y plantas exterior',
        'rangos': [{'tipo': 'prefijos', 'valores': ['8']}]
    },
    'deco_exterior': {
        'descripcion': 'Decoración exterior',
        'rangos': [{'tipo': 'prefijos', 'valores': ['9']}]
    },
    'mascotas_manufacturado': {
        'descripcion': 'Mascotas (productos manufacturados)',
        'rangos': [
            {'tipo': 'prefijos', 'valores': ['2'], 'excluir': CODIGOS_MASCOTAS_VIVO}
        ]
    },
    'mascotas_vivo': {
        'descripcion': 'Mascotas (animales vivos)',
        'rangos': [
            {'tipo': 'codigos_exactos', 'valores': CODIGOS_MASCOTAS_VIVO}
        ]
    },
    'tierra_aridos': {
        'descripcion': 'Tierras y áridos',
        'rangos': [
            {'tipo': 'prefijos', 'valores': ['31', '32']}
        ]
    },
    'fitos': {
        'descripcion': 'Fitosanitarios y abonos',
        'rangos': [
            {'tipo': 'rango', 'valores': ['33', '34', '35', '36', '37', '38', '39']}
        ]
    }
}

# ============================================
# FUNCIONES DE DETECCIÓN DE SECCIÓN
# ============================================

def determinar_seccion(codigo_articulo):
    """
    Determina la sección de un artículo según su código.
    Copiado exactamente de pedido_compras.py
    
    Args:
        codigo_articulo: Código del artículo (puede ser string o número)
    
    Returns:
        str: Nombre de la sección o None si no se puede clasificar
    """
    if codigo_articulo is None:
        return None
    
    codigo_str = str(codigo_articulo).strip()
    
    # Eliminar decimales si viene como float
    if codigo_str.endswith('.0'):
        codigo_str = codigo_str[:-2]
    
    if not codigo_str or codigo_str == 'nan':
        return None
    
    # REGLA CRÍTICA: Filtrar artículos con menos de 10 dígitos
    # Esta regla tiene prioridad sobre todas las demás
    if len(codigo_str) < 10:
        return None
    
    # 1. Verificar códigos de mascotas vivos (primero, tienen prioridad)
    # Los códigos de mascotas vivos son códigos de 4 dígitos (2104, 2204, etc.)
    if codigo_str.startswith('2') and codigo_str[:4] in CODIGOS_MASCOTAS_VIVO:
        return 'mascotas_vivo'
    
    # 2. Sección 2: Mascotas manufacturadas (empieza por 2 y no está en vivos)
    if codigo_str.startswith('2'):
        return 'mascotas_manufacturado'
    
    # 3. Sección 3: Tierra/Áridos (31 o 32)
    if codigo_str.startswith('31') or codigo_str.startswith('32'):
        return 'tierra_aridos'
    
    # 4. Sección 3: Fitosanitarios (33-39)
    if codigo_str.startswith('3'):
        if len(codigo_str) >= 2:
            try:
                segundo_digito = int(codigo_str[1])
                if 3 <= segundo_digito <= 9:
                    return 'fitos'
            except (ValueError, IndexError):
                pass
    
    # 5. Secciones por primer dígito
    if codigo_str.startswith('1'):
        return 'interior'
    elif codigo_str.startswith('4'):
        return 'utiles_jardin'
    elif codigo_str.startswith('5'):
        return 'semillas'
    elif codigo_str.startswith('6'):
        return 'deco_interior'
    elif codigo_str.startswith('7'):
        return 'maf'
    elif codigo_str.startswith('8'):
        return 'vivero'
    elif codigo_str.startswith('9'):
        return 'deco_exterior'
    
    return None

def obtener_descripcion_seccion(seccion):
    """
    Obtiene la descripción legible de una sección.
    
    Args:
        seccion: Nombre de la sección
        
    Returns:
        str: Descripción de la sección o cadena vacía si no existe
    """
    if seccion == 'GENERAL':
        return ''
    return SECCIONES.get(seccion, {}).get('descripcion', '')

# ============================================
# FUNCIONES DE CARGA DE DATOS
# ============================================

def verificar_archivos_seccion(seccion_key):
    """
    Verifica si existen los archivos de una sección específica.
    
    Args:
        seccion_key: Clave de la sección (interior, vivero, etc.)
        
    Returns:
        tuple: (existe_pedido, existe_resumen, pedido_path, resumen_path)
    """
    config = SECTIONS_CONFIG.get(seccion_key, {})
    nombre_archivo = config.get('nombre_archivo', seccion_key)
    
    pedido_path = f'Pedido_compras_{nombre_archivo}.xlsx'
    resumen_path = f'Resumen_Pedidos_{nombre_archivo}.xlsx'
    
    existe_pedido = os.path.exists(pedido_path)
    existe_resumen = os.path.exists(resumen_path)
    
    return existe_pedido, existe_resumen, pedido_path, resumen_path

def cargar_datos_seccion(pedido_xlsx, resumen_xlsx):
    """
    Carga los datos de los archivos Excel de una sección.
    
    Args:
        pedido_xlsx: Ruta al archivo Pedido_compras
        resumen_xlsx: Ruta al archivo Resumen_Pedidos
        
    Returns:
        tuple: (resumen_df, df_todos, top_unidades, top_importe) o (None, None, None, None) si falla
    """
    try:
        # Cargar datos del Resumen_Pedidos
        resumen_df = pd.read_excel(resumen_xlsx, skiprows=1)
        
        # Cargar datos del Pedido_compras para tops
        df_todos = None
        top_unidades = pd.DataFrame()
        top_importe = pd.DataFrame()
        
        try:
            xl = pd.ExcelFile(pedido_xlsx)
            todos_articulos = []
            for sheet in xl.sheet_names:
                if sheet.startswith('Semana_'):
                    df = pd.read_excel(pedido_xlsx, sheet_name=sheet)
                    df['Semana'] = int(sheet.split('_')[1])
                    todos_articulos.append(df)
            
            if todos_articulos:
                df_todos = pd.concat(todos_articulos, ignore_index=True)
                df_filtrado = df_todos[df_todos['Unidades Pedido'] > 0]
                
                # Top por unidades
                top_unidades = df_filtrado.groupby(['Código artículo', 'Nombre Articulo', 'Talla', 'Color'])['Unidades Pedido'].sum().reset_index()
                top_unidades = top_unidades.sort_values('Unidades Pedido', ascending=False).head(10)
                
                # Top por importe
                top_importe = df_filtrado.groupby(['Código artículo', 'Nombre Articulo', 'Talla', 'Color'])['Ventas Objetivo'].sum().reset_index()
                top_importe = top_importe.sort_values('Ventas Objetivo', ascending=False).head(10)
        except Exception as e:
            print(f"      ⚠ No se pudieron procesar tops del pedido: {e}")
        
        return resumen_df, df_todos, top_unidades, top_importe
        
    except Exception as e:
        print(f"      ⚠ Error al cargar datos: {e}")
        return None, None, None, None

# ============================================
# FUNCIONES DE GENERACIÓN DE HTML
# ============================================

def generar_html_seccion(resumen_df, top_unidades, top_importe, seccion_key, es_consolidado=False):
    """
    Genera el contenido HTML para un informe de sección.
    
    Args:
        resumen_df: DataFrame con los datos del resumen
        top_unidades: DataFrame con top 10 por unidades
        top_importe: DataFrame con top 10 por importe
        seccion_key: Clave de la sección
        es_consolidado: Si es True, indica que es el informe consolidado
        
    Returns:
        str: Contenido HTML completo
    """
    config = SECTIONS_CONFIG.get(seccion_key, {})
    
    # Configuración de incrementos festivos
    incrementos_festivos = {
        14: 25,  # Semana Santa
    }
    
    # Preparar datos del resumen
    resumen_semanal = resumen_df.copy()
    resumen_semanal['Semana'] = resumen_semanal['Semana'].astype(int)
    resumen_semanal = resumen_semanal.sort_values('Semana').reset_index(drop=True)
    
    # Calcular métricas globales
    total_semanas = len(resumen_semanal)
    total_articulos = int(resumen_semanal['Total Articulos'].sum())
    total_unidades = int(resumen_semanal['Total Unidades'].sum())
    total_objetivo_importe = float(resumen_semanal['Obj. semana + % crec. + Festivos'].sum())
    stock_minimo_total = int(resumen_semanal['Stock Min Obj'].sum())
    
    # Datos para gráficos
    semanas = resumen_semanal['Semana'].tolist()
    objetivos_importe = resumen_semanal['Obj. semana + % crec. + Festivos'].tolist()
    
    # Generar datos para gráfico SVG
    max_unidades = max(objetivos_importe) if objetivos_importe else 1
    min_unidades = min(objetivos_importe) if objetivos_importe else 0
    
    width = 600
    height = 200
    padding = 40
    
    def get_x(index, total):
        return padding + (index * (width - 2 * padding)) / (total - 1) if total > 1 else width / 2
    
    def get_y(value, max_val, min_val):
        if max_val == min_val:
            return height / 2
        return height - padding - ((value - min_val) / (max_val - min_val)) * (height - 2 * padding)
    
    # Generar path
    points = []
    for i, val in enumerate(objetivos_importe):
        x = get_x(i, len(objetivos_importe))
        y = get_y(val, max_unidades, min_unidades)
        points.append(f"{x:.1f},{y:.1f}")
    
    line_path = " ".join(points)
    area_path = f"{padding},{height - padding} {line_path} {width - padding},{height - padding}"
    
    # Generar week cards HTML
    week_cards_html = ""
    for _, row in resumen_semanal.iterrows():
        semana = int(row['Semana'])
        festivo_value = row['% Festivo'] if '% Festivo' in row.index else incrementos_festivos.get(semana, 0)
        festive_class = "festive" if festivo_value > 0 else ""
        festive_symbol = "+" if festivo_value > 0 else ""
        importe_objetivo = float(row['Obj. semana + % crec. + Festivos'])
        
        week_cards_html += f'''
                <div class="week-card">
                    <div class="week-header">
                        <span class="week-title">Semana {semana}</span>
                        <span class="week-badge {festive_class}">{festive_symbol}{int(festivo_value)}% Festivo</span>
                    </div>
                    <div class="week-metrics">
                        <div class="week-metric">
                            <div class="value">{int(row['Total Unidades']):,}</div>
                            <div class="label">Unidades</div>
                        </div>
                        <div class="week-metric">
                            <div class="value">{int(row['Total Articulos'])}</div>
                            <div class="label">Artículos</div>
                        </div>
                        <div class="week-metric">
                            <div class="value">€{importe_objetivo:,.2f}</div>
                            <div class="label">Importe Obj.</div>
                        </div>
                        <div class="week-metric">
                            <div class="value">{int(row['Stock Min Obj']):,}</div>
                            <div class="label">Stock Mín.</div>
                        </div>
                    </div>
                </div>
'''
    
    # Generar tabla top unidades
    top_unidades_html = ""
    if len(top_unidades) > 0:
        for i, (_, row) in enumerate(top_unidades.iterrows()):
            talla = row['Talla'] if pd.notna(row['Talla']) else ''
            color = row['Color'] if pd.notna(row['Color']) else ''
            top_unidades_html += f'''
                        <tr>
                            <td>{i+1}</td>
                            <td>{row['Código artículo']}</td>
                            <td>{row['Nombre Articulo'][:40]}</td>
                            <td>{talla}</td>
                            <td>{color}</td>
                            <td class="number">{int(row['Unidades Pedido']):,}</td>
                        </tr>
'''
    else:
        top_unidades_html = '''
                        <tr>
                            <td colspan="6" style="text-align:center; color:#666;">No hay datos disponibles</td>
                        </tr>
'''
    
    # Generar tabla top importe
    top_importe_html = ""
    if len(top_importe) > 0:
        for i, (_, row) in enumerate(top_importe.iterrows()):
            talla = row['Talla'] if pd.notna(row['Talla']) else ''
            color = row['Color'] if pd.notna(row['Color']) else ''
            top_importe_html += f'''
                        <tr>
                            <td>{i+1}</td>
                            <td>{row['Código artículo']}</td>
                            <td>{row['Nombre Articulo'][:40]}</td>
                            <td>{talla}</td>
                            <td>{color}</td>
                            <td class="number">€{row['Ventas Objetivo']:,.2f}</td>
                        </tr>
'''
    else:
        top_importe_html = '''
                        <tr>
                            <td colspan="6" style="text-align:center; color:#666;">No hay datos disponibles</td>
                        </tr>
'''
    
    # Generar tabla resumen
    resumen_html = ""
    for _, row in resumen_semanal.iterrows():
        semana = int(row['Semana'])
        festivo_value = row['% Festivo'] if '% Festivo' in row.index else incrementos_festivos.get(semana, 0)
        festive_symbol = "+" if festivo_value > 0 else ""
        importe_objetivo = float(row['Obj. semana + % crec. + Festivos'])
        
        resumen_html += f'''
                        <tr>
                            <td><strong>Semana {semana}</strong></td>
                            <td>{int(row['Total Articulos'])}</td>
                            <td class="number">{int(row['Total Unidades']):,}</td>
                            <td class="number">€{importe_objetivo:,.2f}</td>
                            <td class="number">{int(row['Stock Min Obj']):,}</td>
                            <td class="number">{festive_symbol}{int(festivo_value)}%</td>
                        </tr>
'''
    
    # Generar puntos SVG y labels
    svg_points = ""
    for i in range(len(objetivos_importe)):
        x = get_x(i, len(objetivos_importe))
        y = get_y(objetivos_importe[i], max_unidades, min_unidades)
        svg_points += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5" fill="#1a5f2a" stroke="white" stroke-width="2"/>'
    
    x_labels = ""
    for i in range(len(semanas)):
        x = get_x(i, len(semanas))
        x_labels += f'<text x="{x:.1f}" y="{height-padding+20}" text-anchor="middle" font-size="10" fill="#6c757d">S{semanas[i]}</text>'
    
    # Construir títulos
    if es_consolidado:
        titulo_informe = "INFORME EJECUTIVO CONSOLIDADO - TODAS LAS SECCIONES"
        subtitulo_seccion = "Resumen global de pedidos de compra"
        seccion_display = ""
    else:
        titulo_informe = f"INFORME EJECUTIVO - {config.get('titulo_seccion', seccion_key.upper())}"
        subtitulo_seccion = config.get('descripcion', '')
        seccion_display = subtitulo_seccion
    
    # Construir HTML completo
    html_content = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{titulo_informe} - Vivero Aranjuez 2025</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8eb 100%);
            color: #2c3e50;
            line-height: 1.6;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        header {{
            background: linear-gradient(135deg, #1a5f2a 0%, #28a745 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        header h1 {{
            font-size: 2.2em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}
        
        header .subtitle {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        header .seccion-detectada {{
            font-size: 1.2em;
            margin-top: 10px;
            padding: 8px 20px;
            background: rgba(255,255,255,0.2);
            border-radius: 20px;
            display: inline-block;
        }}
        
        header .date {{
            margin-top: 15px;
            font-size: 0.9em;
            opacity: 0.8;
        }}
        
        .content {{
            padding: 40px;
        }}
        
        .section {{
            margin-bottom: 40px;
        }}
        
        .section h2 {{
            color: #1a5f2a;
            font-size: 1.4em;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #28a745;
        }}
        
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .metric-card {{
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border-radius: 12px;
            padding: 25px;
            text-align: center;
            border: 2px solid #28a745;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(40, 167, 69, 0.3);
        }}
        
        .metric-card .icon {{
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .metric-card .value {{
            font-size: 2em;
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        .metric-card .label {{
            font-size: 0.9em;
            color: #495057;
            margin-top: 5px;
        }}
        
        .chart-container {{
            background: #fafafa;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 30px;
            border: 1px solid #dee2e6;
        }}
        
        .chart-title {{
            font-size: 1.1em;
            color: #1a5f2a;
            margin-bottom: 20px;
            font-weight: 600;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 0.95em;
        }}
        
        .data-table th {{
            background: #1a5f2a;
            color: white;
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
        }}
        
        .data-table td {{
            padding: 12px;
            border: 1px solid #dee2e6;
        }}
        
        .data-table tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        
        .data-table tr:hover {{
            background: #e8f5e9;
        }}
        
        .data-table .number {{
            text-align: right;
            font-weight: 500;
        }}
        
        .week-card {{
            background: #fafafa;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 15px;
            border-left: 4px solid #28a745;
        }}
        
        .week-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }}
        
        .week-title {{
            font-size: 1.2em;
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        .week-badge {{
            background: #28a745;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.85em;
        }}
        
        .week-badge.festive {{
            background: #dc3545;
        }}
        
        .week-metrics {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            gap: 15px;
        }}
        
        .week-metric {{
            text-align: center;
            padding: 10px;
            background: white;
            border-radius: 8px;
        }}
        
        .week-metric .value {{
            font-size: 1.4em;
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        .week-metric .label {{
            font-size: 0.8em;
            color: #6c757d;
        }}
        
        .params-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
        }}
        
        .param-card {{
            background: #f8f9fa;
            border-radius: 10px;
            padding: 20px;
            border: 1px solid #dee2e6;
        }}
        
        .param-card h4 {{
            color: #1a5f2a;
            margin-bottom: 15px;
            font-size: 1.1em;
        }}
        
        .param-item {{
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #e9ecef;
        }}
        
        .param-item:last-child {{
            border-bottom: none;
        }}
        
        .param-value {{
            font-weight: bold;
            color: #28a745;
        }}
        
        .methodology-box {{
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            border-radius: 12px;
            padding: 25px;
            border-left: 4px solid #28a745;
        }}
        
        .methodology-box h4 {{
            color: #1a5f2a;
            margin-bottom: 15px;
        }}
        
        .methodology-box ul {{
            margin-left: 20px;
        }}
        
        .methodology-box li {{
            margin-bottom: 10px;
        }}
        
        .formula-box {{
            background: #f8f9fa;
            border-radius: 8px;
            padding: 15px;
            font-family: 'Courier New', monospace;
            margin: 15px 0;
            border: 1px solid #dee2e6;
        }}
        
        .sections-summary {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}
        
        .section-summary-card {{
            background: #f8f9fa;
            border-radius: 10px;
            padding: 15px;
            border: 1px solid #dee2e6;
            text-align: center;
        }}
        
        .section-summary-card .section-name {{
            font-weight: bold;
            color: #1a5f2a;
            margin-bottom: 5px;
        }}
        
        .section-summary-card .section-value {{
            font-size: 1.5em;
            font-weight: bold;
            color: #28a745;
        }}
        
        footer {{
            background: #1a5f2a;
            color: white;
            text-align: center;
            padding: 20px;
            margin-top: 40px;
        }}
        
        footer p {{
            opacity: 0.9;
            font-size: 0.9em;
        }}
        
        @media (max-width: 768px) {{
            .content {{
                padding: 20px;
            }}
            
            header h1 {{
                font-size: 1.6em;
            }}
            
            .metrics-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>VIVERO ARANJUEZ</h1>
            <p class="subtitle">{titulo_informe}</p>
            {"<p class='seccion-detectada'>" + seccion_display + "</p>" if seccion_display else ""}
            <p class="date">Periodo: Marzo - Mayo 2025 | Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </header>
        
        <div class="content">
            <div class="section">
                <h2>Resumen Ejecutivo</h2>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="icon">📅</div>
                        <div class="value">{total_semanas}</div>
                        <div class="label">Semanas Analizadas</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">📦</div>
                        <div class="value">{total_articulos}</div>
                        <div class="label">Total Articulos</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">🔢</div>
                        <div class="value">{total_unidades:,}</div>
                        <div class="label">Unidades Pedidas</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">💰</div>
                        <div class="value">€{total_objetivo_importe:,.2f}</div>
                        <div class="label">Objetivo Total (+%crec+%fest)</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">🛡️</div>
                        <div class="value">{stock_minimo_total:,}</div>
                        <div class="label">Stock Minimo Total</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>Evolucion Semanal de Pedidos</h2>
                <div class="chart-container">
                    <p class="chart-title">Unidades Pedidas por Semana</p>
                    <svg width="100%" height="250" viewBox="0 0 {width} {height}" preserveAspectRatio="xMidYMid meet">
                        <line x1="{padding}" y1="{height-padding}" x2="{width-padding}" y2="{height-padding}" stroke="#dee2e6" stroke-width="1"/>
                        <line x1="{padding}" y1="{padding}" x2="{padding}" y2="{height-padding}" stroke="#dee2e6" stroke-width="1"/>
                        <text x="{padding-5}" y="{padding+5}" text-anchor="end" font-size="10" fill="#6c757d">{max_unidades:,.0f}</text>
                        <text x="{padding-5}" y="{height-padding}" text-anchor="end" font-size="10" fill="#6c757d">{min_unidades:,.0f}</text>
                        {x_labels}
                        <polygon points="{area_path}" fill="rgba(40, 167, 69, 0.1)" stroke="none"/>
                        <polyline points="{line_path}" fill="none" stroke="#28a745" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
                        {svg_points}
                    </svg>
                </div>
            </div>
            
            <div class="section">
                <h2>Desglose Semanal</h2>
                {week_cards_html}
            </div>
            
            <div class="section">
                <h2>Top 10 Articulos por Volumen</h2>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Codigo</th>
                            <th>Nombre Articulo</th>
                            <th>Talla</th>
                            <th>Color</th>
                            <th class="number">Unidades</th>
                        </tr>
                    </thead>
                    <tbody>
                        {top_unidades_html}
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>Top 10 Articulos por Importe</h2>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Codigo</th>
                            <th>Nombre Articulo</th>
                            <th>Talla</th>
                            <th>Color</th>
                            <th class="number">Importe (EUR)</th>
                        </tr>
                    </thead>
                    <tbody>
                        {top_importe_html}
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>Parametros Aplicados</h2>
                <div class="params-grid">
                    <div class="param-card">
                        <h4>Objetivos de Crecimiento</h4>
                        <div class="param-item">
                            <span>Crecimiento de Ventas</span>
                            <span class="param-value">+5%</span>
                        </div>
                    </div>
                    <div class="param-card">
                        <h4>Stock Minimo Dinamico</h4>
                        <div class="param-item">
                            <span>Porcentaje de Stock</span>
                            <span class="param-value">30%</span>
                        </div>
                    </div>
                    <div class="param-card">
                        <h4>Dias Festivos</h4>
                        <div class="param-item">
                            <span>Semana 14 (Semana Santa)</span>
                            <span class="param-value">+25%</span>
                        </div>
                        <div class="param-item">
                            <span>Semana 18 (1 Mayo)</span>
                            <span class="param-value">Pendiente</span>
                        </div>
                        <div class="param-item">
                            <span>Semana 22 (Festivo Local)</span>
                            <span class="param-value">Pendiente</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>Metodologia de Stock Minimo Dinamico</h2>
                <div class="methodology-box">
                    <h4>Sistema de Calculo Aplicado</h4>
                    <p>El sistema de stock minimo dinamico asegura que cada articulo tenga un nivel de seguridad optimo en tienda, calculando el stock minimo como un porcentaje de las ventas proyectadas de la semana.</p>
                    
                    <div class="formula-box">
                        <strong>Formula de Calculo:</strong><br><br>
                        Stock_Minimo_Objetivo = Ventas_Calculadas × 0.30<br><br>
                        Diferencia_de_Stock = Stock_Minimo_Objetivo - Stock_Acumulado_Anterior<br><br>
                        <strong>Unidades_a_Pedir = Ventas_Calculadas + Diferencia_de_Stock</strong>
                    </div>
                    
                    <h4 style="margin-top: 20px;">Ventajas del Sistema:</h4>
                    <ul>
                        <li><strong>Optimizacion del Capital:</strong> Se utiliza el excedente de semanas anteriores para reducir pedidos actuales</li>
                        <li><strong>Flexibilidad:</strong> El porcentaje de stock minimo se puede ajustar globalmente</li>
                        <li><strong>Seguridad:</strong> Siempre se mantiene un buffer de seguridad proporcional a las ventas</li>
                        <li><strong>Reducciones Inteligentes:</strong> Permite comprar MENOS de las ventas cuando hay excedente de stock</li>
                    </ul>
                    
                    <h4 style="margin-top: 20px;">Ejemplo Practico:</h4>
                    <ul>
                        <li>Semana 13: Ventas calculadas = 90 uds, Stock minimo = 27 uds (30%)</li>
                        <li>Stock acumulado anterior = 30 uds (excedente)</li>
                        <li>Diferencia = 27 - 30 = -3 uds (tenemos 3 uds mas)</li>
                        <li><strong>Pedido = 90 + (-3) = 87 uds</strong> (ahorro por excedente)</li>
                    </ul>
                </div>
            </div>
            
            <div class="section">
                <h2>Resumen Consolidado por Semana</h2>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Semana</th>
                            <th>Articulos</th>
                            <th class="number">Unidades</th>
                            <th class="number">Importe (EUR)</th>
                            <th class="number">Stock Minimo</th>
                            <th class="number">Festivo</th>
                        </tr>
                    </thead>
                    <tbody>
                        {resumen_html}
                    </tbody>
                    <tfoot>
                        <tr style="background: #1a5f2a; color: white; font-weight: bold;">
                            <td>TOTAL</td>
                            <td>{total_articulos}</td>
                            <td class="number">{total_unidades:,}</td>
                            <td class="number">EUR{total_objetivo_importe:,.2f}</td>
                            <td class="number">{stock_minimo_total:,}</td>
                            <td class="number">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
        
        <footer>
            <p><strong>VIVERO ARANJUEZ</strong> - Sistema de Pedidos de Compra 2025</p>
            <p>Informe generado automaticamente con metodologia de Stock Minimo Dinamico</p>
            <p style="margin-top: 10px; font-size: 0.8em;">© {datetime.now().year} Vivero Aranjuez - Todos los derechos reservados</p>
        </footer>
    </div>
</body>
</html>'''
    
    return html_content

def guardar_html(html_content, nombre_archivo):
    """
    Guarda el contenido HTML en un archivo.
    
    Args:
        html_content: Contenido HTML a guardar
        nombre_archivo: Nombre del archivo de salida
        
    Returns:
        str: Ruta del archivo guardado
    """
    with open(nombre_archivo, 'w', encoding='utf-8') as f:
        f.write(html_content)
    return nombre_archivo

# ============================================
# FUNCIONES PRINCIPALES DE GENERACIÓN
# ============================================

def generar_informe_seccion(seccion_key):
    """
    Genera un informe HTML para una sección específica.
    
    Args:
        seccion_key: Clave de la sección (interior, vivero, etc.)
        
    Returns:
        tuple: (exito, archivo_generado, metricas)
    """
    config = SECTIONS_CONFIG.get(seccion_key, {})
    nombre_archivo = config.get('nombre_archivo', seccion_key)
    
    print(f"\n📊 Procesando sección: {config.get('titulo_seccion', seccion_key)}")
    
    # Verificar archivos
    existe_pedido, existe_resumen, pedido_path, resumen_path = verificar_archivos_seccion(seccion_key)
    
    if not existe_resumen:
        print(f"   ⚠ No encontrado: {resumen_path}")
        return False, None, None
    
    # Cargar datos
    resumen_df, df_todos, top_unidades, top_importe = cargar_datos_seccion(
        pedido_path if existe_pedido else None,
        resumen_path
    )
    
    if resumen_df is None:
        return False, None, None
    
    # Generar HTML
    html_content = generar_html_seccion(resumen_df, top_unidades, top_importe, seccion_key)
    
    # Guardar archivo
    output_path = f'Informe_Compra_{nombre_archivo}.html'
    guardar_html(html_content, output_path)
    
    # Calcular métricas
    metricas = {
        'seccion': config.get('titulo_seccion', seccion_key),
        'total_articulos': int(resumen_df['Total Articulos'].sum()),
        'total_unidades': int(resumen_df['Total Unidades'].sum()),
        'total_importe': float(resumen_df['Obj. semana + % crec. + Festivos'].sum())
    }
    
    print(f"   ✅ Generado: {output_path}")
    print(f"      - Artículos: {metricas['total_articulos']:,}")
    print(f"      - Unidades: {metricas['total_unidades']:,}")
    print(f"      - Importe: €{metricas['total_importe']:,.2f}")
    
    return True, output_path, metricas

def generar_informe_consolidado(lista_secciones, todas_metricas):
    """
    Genera un informe consolidado con todas las secciones.
    
    Args:
        lista_secciones: Lista de secciones procesadas exitosamente
        todas_metricas: Lista de métricas de cada sección
        
    Returns:
        str: Ruta del archivo generado
    """
    print(f"\n📋 Generando informe consolidado...")
    
    # Consolidar datos de todas las secciones
    datos_consolidados = []
    
    for seccion_key in lista_secciones:
        config = SECTIONS_CONFIG.get(seccion_key, {})
        nombre_archivo = config.get('nombre_archivo', seccion_key)
        _, _, resumen_path = verificar_archivos_seccion(seccion_key)[2:]
        
        if os.path.exists(resumen_path):
            try:
                resumen_df = pd.read_excel(resumen_path, skiprows=1)
                resumen_df['Seccion'] = config.get('titulo_seccion', seccion_key)
                datos_consolidados.append(resumen_df)
            except:
                pass
    
    if not datos_consolidados:
        print("   ⚠ No hay datos para consolidar")
        return None
    
    # Concatenar todos los datos
    df_consolidado = pd.concat(datos_consolidados, ignore_index=True)
    
    # Calcular métricas globales
    metricas_consolidadas = {
        'total_secciones': len(lista_secciones),
        'total_articulos': int(df_consolidado['Total Articulos'].sum()),
        'total_unidades': int(df_consolidado['Total Unidades'].sum()),
        'total_importe': float(df_consolidado['Obj. semana + % crec. + Festivos'].sum())
    }
    
    # Generar HTML consolidado (simplificado para el resumen global)
    html_content = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe Consolidado - Vivero Aranjuez 2025</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8eb 100%);
            color: #2c3e50;
            line-height: 1.6;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        header {{
            background: linear-gradient(135deg, #1a5f2a 0%, #28a745 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        header h1 {{
            font-size: 2.2em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}
        
        header .subtitle {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        header .date {{
            margin-top: 15px;
            font-size: 0.9em;
            opacity: 0.8;
        }}
        
        .content {{
            padding: 40px;
        }}
        
        .section {{
            margin-bottom: 40px;
        }}
        
        .section h2 {{
            color: #1a5f2a;
            font-size: 1.4em;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #28a745;
        }}
        
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .metric-card {{
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border-radius: 12px;
            padding: 25px;
            text-align: center;
            border: 2px solid #28a745;
        }}
        
        .metric-card .icon {{
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .metric-card .value {{
            font-size: 2em;
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        .metric-card .label {{
            font-size: 0.9em;
            color: #495057;
            margin-top: 5px;
        }}
        
        .sections-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
        }}
        
        .section-card {{
            background: #f8f9fa;
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #dee2e6;
            border-left: 4px solid #28a745;
        }}
        
        .section-card h3 {{
            color: #1a5f2a;
            margin-bottom: 15px;
            font-size: 1.1em;
        }}
        
        .section-metric {{
            display: flex;
            justify-content: space-between;
            padding: 8px 0;
            border-bottom: 1px solid #e9ecef;
        }}
        
        .section-metric:last-child {{
            border-bottom: none;
        }}
        
        .section-metric .label {{
            color: #6c757d;
        }}
        
        .section-metric .value {{
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 0.95em;
        }}
        
        .data-table th {{
            background: #1a5f2a;
            color: white;
            padding: 14px 12px;
            text-align: left;
            font-weight: 600;
        }}
        
        .data-table td {{
            padding: 12px;
            border: 1px solid #dee2e6;
        }}
        
        .data-table tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        
        .data-table tr:hover {{
            background: #e8f5e9;
        }}
        
        .data-table .number {{
            text-align: right;
            font-weight: 500;
        }}
        
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .summary-item {{
            display: flex;
            justify-content: space-between;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            border: 1px solid #dee2e6;
        }}
        
        .summary-item .label {{
            font-weight: 600;
            color: #495057;
        }}
        
        .summary-item .value {{
            font-size: 1.3em;
            font-weight: bold;
            color: #1a5f2a;
        }}
        
        footer {{
            background: #1a5f2a;
            color: white;
            text-align: center;
            padding: 20px;
            margin-top: 40px;
        }}
        
        footer p {{
            opacity: 0.9;
            font-size: 0.9em;
        }}
        
        @media (max-width: 768px) {{
            .content {{
                padding: 20px;
            }}
            
            header h1 {{
                font-size: 1.6em;
            }}
            
            .metrics-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>VIVERO ARANJUEZ</h1>
            <p class="subtitle">INFORME EJECUTIVO CONSOLIDADO - TODAS LAS SECCIONES</p>
            <p class="date">Periodo: Marzo - Mayo 2025 | Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </header>
        
        <div class="content">
            <div class="section">
                <h2>Resumen Global</h2>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="icon">📊</div>
                        <div class="value">{metricas_consolidadas['total_secciones']}</div>
                        <div class="label">Secciones Procesadas</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">📦</div>
                        <div class="value">{metricas_consolidadas['total_articulos']:,}</div>
                        <div class="label">Total Articulos</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">🔢</div>
                        <div class="value">{metricas_consolidadas['total_unidades']:,}</div>
                        <div class="label">Unidades Totales</div>
                    </div>
                    <div class="metric-card">
                        <div class="icon">💰</div>
                        <div class="value">€{metricas_consolidadas['total_importe']:,.2f}</div>
                        <div class="label">Importe Total Objetivo</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>Resumen por Sección</h2>
                <div class="sections-grid">
'''
    
    # Agregar cards de cada sección
    for metrica in todas_metricas:
        html_content += f'''
                    <div class="section-card">
                        <h3>{metrica['seccion']}</h3>
                        <div class="section-metric">
                            <span class="label">Artículos</span>
                            <span class="value">{metrica['total_articulos']:,}</span>
                        </div>
                        <div class="section-metric">
                            <span class="label">Unidades</span>
                            <span class="value">{metrica['total_unidades']:,}</span>
                        </div>
                        <div class="section-metric">
                            <span class="label">Importe</span>
                            <span class="value">€{metrica['total_importe']:,.2f}</span>
                        </div>
                    </div>
'''
    
    html_content += '''
                </div>
            </div>
            
            <div class="section">
                <h2>Detalle por Semana (Todas las Secciones)</h2>
'''
    
    # Generar tabla consolidada
    if df_consolidado is not None and len(df_consolidado) > 0:
        semanas_unicas = sorted(df_consolidado['Semana'].unique())
        
        html_content += '''                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Semana</th>
                            <th>Secciones</th>
                            <th class="number">Total Articulos</th>
                            <th class="number">Total Unidades</th>
                            <th class="number">Total Importe</th>
                        </tr>
                    </thead>
                    <tbody>
'''
        
        for semana in semanas_unicas:
            df_semana = df_consolidado[df_consolidado['Semana'] == semana]
            articulos = int(df_semana['Total Articulos'].sum())
            unidades = int(df_semana['Total Unidades'].sum())
            importe = float(df_semana['Obj. semana + % crec. + Festivos'].sum())
            num_secciones = df_semana['Seccion'].nunique()
            
            html_content += f'''
                        <tr>
                            <td><strong>Semana {int(semana)}</strong></td>
                            <td>{num_secciones} secciones</td>
                            <td class="number">{articulos:,}</td>
                            <td class="number">{unidades:,}</td>
                            <td class="number">€{importe:,.2f}</td>
                        </tr>
'''
        
        html_content += f'''
                    </tbody>
                    <tfoot>
                        <tr style="background: #1a5f2a; color: white; font-weight: bold;">
                            <td>TOTAL</td>
                            <td>{metricas_consolidadas['total_secciones']} secciones</td>
                            <td class="number">{metricas_consolidadas['total_articulos']:,}</td>
                            <td class="number">{metricas_consolidadas['total_unidades']:,}</td>
                            <td class="number">€{metricas_consolidadas['total_importe']:,.2f}</td>
                        </tr>
                    </tfoot>
                </table>
'''
    
    html_content += '''            </div>
            
            <div class="section">
                <h2>Archivos Generados</h2>
'''
    
    # Listar todos los informes generados
    html_content += '''                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Sección</th>
                            <th>Archivo Generado</th>
                        </tr>
                    </thead>
                    <tbody>
'''
    
    for metrica in todas_metricas:
        seccion_key = metrica['seccion'].lower().replace(' ', '_').replace('(', '').replace(')', '')
        nombre_archivo = f"Informe_Compra_{seccion_key}.html"
        html_content += f'''
                        <tr>
                            <td>{metrica['seccion']}</td>
                            <td><a href="{nombre_archivo}">{nombre_archivo}</a></td>
                        </tr>
'''
    
    html_content += '''                    </tbody>
                </table>
            </div>
        </div>
        
        <footer>
            <p><strong>VIVERO ARANJUEZ</strong> - Sistema de Pedidos de Compra 2025</p>
            <p>Informe consolidado generado automaticamente con metodologia de Stock Minimo Dinamico</p>
            <p style="margin-top: 10px; font-size: 0.8em;">© ''' + str(datetime.now().year) + ''' Vivero Aranjuez - Todos los derechos reservados</p>
        </footer>
    </div>
</body>
</html>'''
    
    # Guardar archivo consolidado
    output_path = 'Informe_Consolidado_Todas_Secciones.html'
    guardar_html(html_content, output_path)
    
    print(f"   ✅ Generado: {output_path}")
    print(f"   📊 Total: {metricas_consolidadas['total_secciones']} secciones procesadas")
    
    return output_path

def generar_todos_los_informes():
    """
    Genera informes para todas las secciones disponibles y un informe consolidado.
    
    Returns:
        tuple: (lista_informes, metricas_consolidadas)
    """
    print("=" * 70)
    print("GENERADOR DE INFORMES EJECUTIVOS HTML - MÚLTIPLES SECCIONES")
    print("   Vivero Aranjuez 2025")
    print("   Generación automática de informes por sección + consolidado")
    print("=" * 70)
    
    secciones_procesadas = []
    todas_las_metricas = []
    informes_generados = []
    
    # Iterar por todas las secciones definidas
    for seccion_key in SECTIONS_CONFIG.keys():
        exito, archivo, metricas = generar_informe_seccion(seccion_key)
        if exito:
            secciones_procesadas.append(seccion_key)
            todas_las_metricas.append(metricas)
            informes_generados.append(archivo)
    
    # Generar informe consolidado si hay secciones procesadas
    if secciones_procesadas:
        archivo_consolidado = generar_informe_consolidado(secciones_procesadas, todas_las_metricas)
        if archivo_consolidado:
            informes_generados.append(archivo_consolidado)
    
    # Resumen final
    print("\n" + "=" * 70)
    print("RESUMEN DE GENERACIÓN DE INFORMES")
    print("=" * 70)
    print(f"📊 Secciones procesadas: {len(secciones_procesadas)}/{len(SECTIONS_CONFIG)}")
    print(f"📄 Informes individuales generados: {len(secciones_procesadas)}")
    print(f"📋 Informes totales generados: {len(informes_generados)}")
    
    if informes_generados:
        print(f"\n📁 Archivos generados:")
        for archivo in informes_generados:
            print(f"   - {archivo}")
    
    print("\n✅ Proceso completado exitosamente!")
    
    return informes_generados, todas_las_metricas

def main():
    """Función principal que ejecuta la generación de todos los informes."""
    generar_todos_los_informes()

if __name__ == "__main__":
    main()
