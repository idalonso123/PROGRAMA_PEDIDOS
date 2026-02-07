#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INFORME_FINAL.py
Script para generar el informe HTML completo del análisis ABC+D de Vivero Aranjuez.
Versión mejorada: Lee archivos CLASIFICACION_ABC+D_[SECCION].xlsx con datos ya clasificados.
Envía automáticamente un email con todos los informes generados a Ivan.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import glob
import warnings
import os
import smtplib
import ssl
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from pathlib import Path
warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURACIÓN DE EMAIL
# ============================================================================

# Destinatario: Ivan (información del archivo src/email_service.py)
DESTINATARIO_IVAN = {
    'nombre': 'Ivan',
    'email': 'ivan.delgado@viveverde.es'
}

# Configuración del servidor SMTP
SMTP_CONFIG = {
    'servidor': 'smtp.serviciodecorreo.es',
    'puerto': 465,
    'remitente_email': 'ivan.delgado@viveverde.es',
    'remitente_nombre': 'Sistema de Pedidos automáticos VIVEVERDE'
}

# Configuración de fechas
FECHA_INICIO = datetime(2025, 1, 1)
FECHA_FIN = datetime(2025, 2, 28)
DIAS_PERIODO = 59

# Período formateado para nombres de archivo (formato: YYYYMMDD-YYYYMMDD)
PERIODO_FILENAME = f"{FECHA_INICIO.strftime('%Y%m%d')}-{FECHA_FIN.strftime('%Y%m%d')}"

# Período formateado para el email
PERIODO_EMAIL = f"{FECHA_INICIO.strftime('%d/%m/%Y')} - {FECHA_FIN.strftime('%d/%m/%Y')}"

# ============================================================================
# FUNCIÓN PARA ENVIAR EMAIL CON INFORMES ADJUNTOS
# ============================================================================

def enviar_email_informes(archivos_informes: list) -> bool:
    """
    Envía un email a Ivan con todos los informes HTML generados adjuntos.
    
    Args:
        archivos_informes: Lista de rutas de archivos HTML generados
    
    Returns:
        bool: True si el email fue enviado exitosamente, False en caso contrario
    """
    if not archivos_informes:
        print("  AVISO: No hay informes para enviar. No se enviará email.")
        return False
    
    nombre_destinatario = DESTINATARIO_IVAN['nombre']
    email_destinatario = DESTINATARIO_IVAN['email']
    
    # Verificar que todos los archivos existen
    archivos_existentes = []
    for archivo in archivos_informes:
        if Path(archivo).exists():
            archivos_existentes.append(archivo)
        else:
            print(f"  AVISO: El archivo '{archivo}' no existe y no se adjuntará.")
    
    if not archivos_existentes:
        print("  AVISO: No hay archivos válidos para adjuntar. No se enviará email.")
        return False
    
    # Verificar contraseña en variable de entorno
    password = os.environ.get('EMAIL_PASSWORD')
    if not password:
        print(f"  AVISO: Variable de entorno 'EMAIL_PASSWORD' no configurada. No se enviará email a {nombre_destinatario}.")
        return False
    
    try:
        # Crear mensaje MIME
        msg = MIMEMultipart()
        msg['From'] = f"{SMTP_CONFIG['remitente_nombre']} <{SMTP_CONFIG['remitente_email']}>"
        msg['To'] = email_destinatario
        msg['Subject'] = f"VIVEVERDE: Informes de ClasificacionABC+D de cada sección del periodo {PERIODO_EMAIL}"
        
        # Cuerpo del email
        cuerpo = f"""Buenos días {nombre_destinatario},

Te adjunto en este correo los informes de Clasificación ABC+D de cada sección.

Atentamente,

Sistema de Pedidos automáticos VIVEVERDE."""
        
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar archivos HTML
        for archivo in archivos_existentes:
            try:
                filename = Path(archivo).name
                with open(archivo, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= "{filename}"')
                part.add_header('Content-Type', 'text/html')
                msg.attach(part)
                print(f"  Adjunto añadido: {filename}")
            except Exception as e:
                print(f"  ERROR al adjuntar archivo {archivo}: {e}")
        
        # Enviar email mediante SSL
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_CONFIG['servidor'], SMTP_CONFIG['puerto'], context=context) as server:
            server.login(SMTP_CONFIG['remitente_email'], password)
            server.sendmail(SMTP_CONFIG['remitente_email'], email_destinatario, msg.as_string())
        
        print(f"  Email enviado a {nombre_destinatario} ({email_destinatario})")
        return True
        
    except smtplib.SMTPException as e:
        print(f"  ERROR SMTP al enviar email a {nombre_destinatario}: {e}")
        return False
    except Exception as e:
        print(f"  ERROR al enviar email a {nombre_destinatario}: {e}")
        return False

def obtener_archivos_clasificacion():
    """Busca todos los archivos de clasificación ABC+D por sección.
    
    Soporta tanto el formato antiguo (CLASIFICACION_ABC+D_SECCION.xlsx) como
    el nuevo formato con período y año (CLASIFICACION_ABC+D_SECCION_PERIODO_AÑO.xlsx).
    """
    patrones = [
        "data/input/CLASIFICACION_ABC+D_*.xlsx",
    ]
    
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
    """Extrae el nombre de la sección del nombre del archivo.
    
    Soporta los siguientes formatos:
    - Formato antiguo: CLASIFICACION_ABC+D_INTERIOR.xlsx -> devuelve "INTERIOR"
    - Formato nuevo: CLASIFICACION_ABC+D_INTERIOR_P2_2025.xlsx -> devuelve "INTERIOR"
    - Secciones compuestas: CLASIFICACION_ABC+D_MASCOTAS_MANUFACTURADO_P2_2025.xlsx
      -> devuelve "MASCOTAS_MANUFACTURADO"
    
    Las secciones válidas son: INTERIOR, EXTERIOR, JARDINERIA, MACETAS, COMPLEMENTOS, 
    FITOSANITARIOS, MASCOTAS_MANUFACTURADO, MASCOTAS_VIVO, TIERRAS_ARIDOS
    """
    basename = os.path.basename(nombre_archivo)
    nombre_sin_extension = basename.replace('.xlsx', '')
    prefijo = "CLASIFICACION_ABC+D_"
    
    if not nombre_sin_extension.startswith(prefijo):
        return None
    
    # Extraer la parte después del prefijo
    nombre_sin_prefijo = nombre_sin_extension[len(prefijo):]
    
    # Definir las secciones válidas conocidas (incluyendo secciones compuestas)
    secciones_validas = {
        'INTERIOR', 'EXTERIOR', 'JARDINERIA', 'MACETAS', 'COMPLEMENTOS', 'FITOSANITARIOS',
        'MASCOTAS_MANUFACTURADO', 'MASCOTAS_VIVO', 'TIERRAS_ARIDOS',
        # Versiones en minúsculas para comparación
        'interior', 'exterior', 'jardineria', 'macetas', 'complementos', 'fitosanitarios',
        'mascotas_manufacturado', 'mascotas_vivo', 'tierras_aridos'
    }
    
    # Partes del nombre separadas por guión bajo
    partes = nombre_sin_prefijo.split('_')
    
    if len(partes) >= 3:
        # Intentar reconstruir el nombre de la sección
        # Las secciones pueden ser simples (INTERIOR) o compuestas (MASCOTAS_MANUFACTURADO)
        for num_partes_seccion in range(1, min(4, len(partes) - 1)):
            # Construir nombre de sección candidato usando las primeras N partes
            seccion_candidato = '_'.join(partes[:num_partes_seccion]).upper()
            
            if seccion_candidato in secciones_validas:
                # Verificar que después de la sección viene el período
                # El formato nuevo tiene: SECCION_PERIODO_AÑO o SECCION_PERIODO_AÑOS
                indice_siguiente = num_partes_seccion
                
                # Debe haber al menos: período + año (2 partes mínimo) o período + X + año (3 partes)
                partes_restantes = len(partes) - indice_siguiente
                
                if partes_restantes >= 2:
                    parte_periodo = partes[indice_siguiente]
                    parte_anio = partes[indice_siguiente + partes_restantes - 1]
                    
                    # Verificar formato de período (P1, P2, P3, P4)
                    es_periodo_valido = parte_periodo.upper().startswith('P') and len(parte_periodo) <= 3
                    
                    # Verificar formato de año (4 dígitos)
                    es_anio_valido = parte_anio.isdigit() and len(parte_anio) == 4
                    
                    if es_periodo_valido and es_anio_valido:
                        # Es el formato nuevo, devolver solo la sección
                        return seccion_candidato
    
    # Formato antiguo o no reconocido, devolver todo después del prefijo
    return nombre_sin_prefijo

def leer_datos_clasificacion(ruta_archivo):
    """Lee todas las hojas de clasificación del archivo Excel."""
    excel_file = pd.ExcelFile(ruta_archivo)
    hojas = {}
    for hoja in excel_file.sheet_names:
        hojas[hoja] = pd.read_excel(excel_file, sheet_name=hoja)
    return hojas

def obtener_valor(diccionario, clave, default=0):
    """Obtiene un valor de un diccionario o Serie de forma segura."""
    try:
        val = diccionario[clave]
        if pd.isna(val):
            return default
        if isinstance(val, (np.integer, np.floating)):
            return int(val) if isinstance(val, np.integer) else float(val)
        return val
    except:
        return default

def normalizar_articulo(valor):
    """
    Normaliza un código de artículo para que pueda compararse correctamente.
    Maneja:
    - Formato float (2304030011.0 → 2304030011)
    - Notación científica (1.010100e+08 → 101010001)
    - Valores ya formateados como string
    """
    try:
        # Convertir a string primero para manejar notación científica
        valor_str = str(valor).strip()
        
        # Si está vacío, devolver None
        if not valor_str or valor_str == 'nan':
            return None
        
        # Convertir a float (maneja notación científica y .0)
        valor_float = float(valor_str)
        
        # Convertir a int y luego a string
        return str(int(valor_float))
        
    except Exception:
        return None

def leer_capital_inmovilizado_stock(df_seccion):
    """
    Lee el archivo de stock y calcula el capital inmovilizado real de los artículos
    de la sección específica, sumando la columna 'Total' solo para esos artículos.
    
    Busca archivos con el patrón SPA_stock_P*.xlsx (por ejemplo: SPA_stock_P1.xlsx,
    SPA_stock_P2.xlsx, etc.). Usa el archivo más reciente disponible.
    
    Args:
        df_seccion: DataFrame con los artículos de la sección (del archivo CLASIFICACION_ABC+D)
    
    Returns:
        float: Capital inmovilizado total para los artículos de la sección
    """
    # Buscar archivos con el patrón SPA_stock_P*.xlsx
    patrones_stock = [
        "data/input/SPA_stock_P*.xlsx",
        "data/input/stock.xlsx"  # Fallback legacy
    ]
    
    archivo_encontrado = None
    for patron in patrones_stock:
        archivos = glob.glob(patron)
        if archivos:
            # Ordenar por nombre (P4 > P3 > P2 > P1)
            archivos.sort(reverse=True)
            archivo_encontrado = archivos[0]
            break
    
    if archivo_encontrado:
        print(f"    ✓ Archivo de stock encontrado: {archivo_encontrado}")
    else:
        print(f"    Advertencia: No se encontró ningún archivo de stock")
        print(f"      - Patrones buscados: {patrones_stock}")
        return None
    
    ruta_stock = archivo_encontrado
    
    try:
        # Leer stock, excluyendo filas Cabecera (sumatorios)
        df_stock = pd.read_excel(ruta_stock)
        df_stock = df_stock[df_stock['Tipo registro'] != 'Cabecera']
        
        if 'Total' not in df_stock.columns:
            print(f"    Advertencia: No se encontró la columna 'Total' en el archivo de stock")
            return None
        
        # Forward-fill: Rellenar celdas vacías de Artículo y Nombre artículo
        # Solo para filas Detalle (forward-fill propagation)
        df_stock['Artículo'] = df_stock['Artículo'].ffill()
        df_stock['Nombre artículo'] = df_stock['Nombre artículo'].ffill()
        
        # Normalizar artículos en el stock
        df_stock['Artículo'] = df_stock['Artículo'].apply(normalizar_articulo)
        
        # Normalizar artículos en la sección
        df_seccion['Artículo'] = df_seccion['Artículo'].apply(normalizar_articulo)
        
        # Eliminar None values (artículos que no se pudieron normalizar)
        df_stock = df_stock[df_stock['Artículo'].notna()]
        df_seccion = df_seccion[df_seccion['Artículo'].notna()]
        
        # Obtener artículos únicos de la sección
        articulos_seccion = set(df_seccion['Artículo'].unique())
        
        # Filtrar stock por los artículos de la sección
        df_filtrado = df_stock[df_stock['Artículo'].isin(articulos_seccion)]
        
        # Sumar la columna Total
        capital_inmovilizado = df_filtrado['Total'].sum()
        
        # Manejar posibles NaN
        if pd.isna(capital_inmovilizado):
            capital_inmovilizado = 0
        
        print(f"    ✓ Capital inmovilizado leído del stock: {capital_inmovilizado:,.2f}€ ({len(df_filtrado)} filas sumadas)")
        
        return capital_inmovilizado
        
    except Exception as e:
        print(f"    Error al leer el capital inmovilizado del stock: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def formatear_numero(valor, decimales=0):
    """Formatea un número con separadores de miles."""
    if valor is None:
        return "0"
    try:
        return f"{float(valor):,.{decimales}f}"
    except:
        return "0"

def calcular_nivel_stock(stock):
    """Calcula el nivel de stock basándose en las unidades."""
    if pd.isna(stock) or stock == 0:
        return 'CERO'
    elif stock <= 5:
        return 'BAJO'
    elif stock <= 20:
        return 'NORMAL'
    else:
        return 'ELEVADO'

def calcular_nivel_riesgo(rotacion):
    """
    Calcula el nivel de riesgo basándose en el % de Rotación Consumido.
    
    Umbrales:
    - CRITICO: > 150% (superado significativamente el período óptimo)
    - ALTO: 100-150% (superado el período óptimo)
    - MEDIO: 65-100% (cerca del límite)
    - BAJO: < 65% (período óptimo de rotación)
    """
    if pd.isna(rotacion) or rotacion == 0:
        return 'BAJO'
    elif rotacion > 150:
        return 'CRITICO'
    elif rotacion >= 100:
        return 'ALTO'
    elif rotacion >= 65:
        return 'MEDIO'
    else:
        return 'BAJO'

def normalizar_riesgo(valor):
    """Normaliza el valor de riesgo a mayúsculas para comparación."""
    if pd.isna(valor):
        return 'BAJO'
    return str(valor).upper().strip()

def generar_html_informe(datos, df_completo, nombre_seccion=None):
    """Genera el HTML completo del informe."""
    r = datos['resumen']
    dc = datos['dist_categoria']
    ds = datos['dist_stock']
    dri = datos['dist_riesgo']
    tv = datos['top_ventas']
    tr = datos['top_riesgo']
    te = datos['top_estrella']
    
    fecha_actual = datetime.now().strftime("%d de %B de %Y")
    nombre_seccion_titulo = nombre_seccion.upper() if nombre_seccion else "Vivero"
    
    # Obtener valores por categoría
    cat_a = dc[dc['categoria'] == 'A'].iloc[0] if 'A' in dc['categoria'].values else None
    cat_b = dc[dc['categoria'] == 'B'].iloc[0] if 'B' in dc['categoria'].values else None
    cat_c = dc[dc['categoria'] == 'C'].iloc[0] if 'C' in dc['categoria'].values else None
    cat_d = dc[dc['categoria'] == 'D'].iloc[0] if 'D' in dc['categoria'].values else None
    
    count_a = obtener_valor(cat_a, 'articulos', 0)
    count_b = obtener_valor(cat_b, 'articulos', 0)
    count_c = obtener_valor(cat_c, 'articulos', 0)
    count_d = obtener_valor(cat_d, 'articulos', 0)
    
    ventas_a = obtener_valor(cat_a, 'ventas', 0)
    ventas_b = obtener_valor(cat_b, 'ventas', 0)
    ventas_c = obtener_valor(cat_c, 'ventas', 0)
    stock_a = obtener_valor(cat_a, 'stock', 0)
    stock_b = obtener_valor(cat_b, 'stock', 0)
    stock_c = obtener_valor(cat_c, 'stock', 0)
    stock_d = obtener_valor(cat_d, 'stock', 0)
    
    total_arts = count_a + count_b + count_c + count_d
    total_ventas = ventas_a + ventas_b + ventas_c
    
    pct_a = round(count_a / total_arts * 100, 1) if total_arts > 0 else 0
    pct_b = round(count_b / total_arts * 100, 1) if total_arts > 0 else 0
    pct_c = round(count_c / total_arts * 100, 1) if total_arts > 0 else 0
    pct_d = round(count_d / total_arts * 100, 1) if total_arts > 0 else 0
    
    pct_ventas_a = round(ventas_a / total_ventas * 100, 1) if total_ventas > 0 else 0
    pct_ventas_b = round(ventas_b / total_ventas * 100, 1) if total_ventas > 0 else 0
    pct_ventas_c = round(ventas_c / total_ventas * 100, 1) if total_ventas > 0 else 0
    
    # Datos de riesgo desde el archivo
    riesgo_critico = dri[dri['nivel'] == 'CRITICO']['articulos'].values
    riesgo_alto = dri[dri['nivel'] == 'ALTO']['articulos'].values
    riesgo_medio = dri[dri['nivel'] == 'MEDIO']['articulos'].values
    riesgo_bajo = dri[dri['nivel'] == 'BAJO']['articulos'].values
    
    count_critico = int(riesgo_critico[0]) if len(riesgo_critico) > 0 else 0
    count_alto = int(riesgo_alto[0]) if len(riesgo_alto) > 0 else 0
    count_medio = int(riesgo_medio[0]) if len(riesgo_medio) > 0 else 0
    count_bajo_riesgo = int(riesgo_bajo[0]) if len(riesgo_bajo) > 0 else 0
    
    # Datos de stock desde el archivo
    stock_elevado = ds[ds['nivel'] == 'ELEVADO']['articulos'].values
    stock_normal = ds[ds['nivel'] == 'NORMAL']['articulos'].values
    stock_bajo = ds[ds['nivel'] == 'BAJO']['articulos'].values
    stock_cero = ds[ds['nivel'] == 'CERO']['articulos'].values
    
    count_elevado = int(stock_elevado[0]) if len(stock_elevado) > 0 else 0
    count_normal = int(stock_normal[0]) if len(stock_normal) > 0 else 0
    count_bajo_stock = int(stock_bajo[0]) if len(stock_bajo) > 0 else 0
    count_cero_stock = int(stock_cero[0]) if len(stock_cero) > 0 else 0
    
    pct_elevado = round(count_elevado / total_arts * 100, 1) if total_arts > 0 else 0
    pct_normal = round(count_normal / total_arts * 100, 1) if total_arts > 0 else 0
    pct_bajo_stock = round(count_bajo_stock / total_arts * 100, 1) if total_arts > 0 else 0
    pct_cero_stock = round(count_cero_stock / total_arts * 100, 1) if total_arts > 0 else 0
    
    # Calcular ángulos dinámicos para el diagrama de circunferencia de Antigüedad del Stock
    total_stock_count = count_elevado + count_normal + count_bajo_stock + count_cero_stock
    if total_stock_count > 0:
        # Convertir counts a ángulos (360° = 100%)
        angle_elevado = round(count_elevado / total_stock_count * 360, 1)
        angle_normal = round(count_normal / total_stock_count * 360, 1)
        angle_bajo = round(count_bajo_stock / total_stock_count * 360, 1)
        # El último ángulo se calcula por residuo para completar 360°
        angle_cero = round(360 - angle_elevado - angle_normal - angle_bajo, 1)
        
        # Calcular puntos de corte para el gradiente cónico
        # Cada sector comienza donde termina el anterior
        end_elevado = angle_elevado
        end_normal = end_elevado + angle_normal
        end_bajo = end_normal + angle_bajo
        
        # Generar la cadena del gradiente cónico dinámicamente
        # Colores: ELEVADO=verde(#C8E6C9), NORMAL=amarillo(#FFF9C4), BAJO=naranja(#FFE0B2), CERO=rojo(#FFCDD2)
        chart_gradient = f"#C8E6C9 0deg {end_elevado}deg, #FFF9C4 {end_elevado}deg {end_normal}deg, #FFE0B2 {end_normal}deg {end_bajo}deg, #FFCDD2 {end_bajo}deg 360deg"
    else:
        # Valores por defecto si no hay datos
        chart_gradient = "#C8E6C9 0deg 90deg, #FFF9C4 90deg 180deg, #FFE0B2 180deg 270deg, #FFCDD2 270deg 360deg"
        end_elevado = 90
        end_normal = 180
        end_bajo = 270
    
    # Calcular valores para stock
    stock_elevado_sum = int(df_completo[df_completo['nivel_stock'] == 'ELEVADO']['Stock Final (unidades)'].sum()) if 'Stock Final (unidades)' in df_completo.columns else 0
    stock_normal_sum = int(df_completo[df_completo['nivel_stock'] == 'NORMAL']['Stock Final (unidades)'].sum()) if 'Stock Final (unidades)' in df_completo.columns else 0
    stock_bajo_sum = int(df_completo[df_completo['nivel_stock'] == 'BAJO']['Stock Final (unidades)'].sum()) if 'Stock Final (unidades)' in df_completo.columns else 0
    
    # Calcular matriz cruzando nivel_stock con nivel_riesgo
    matrix = {}
    niveles_stock = ['ELEVADO', 'NORMAL', 'BAJO', 'CERO']
    niveles_riesgo = ['BAJO', 'MEDIO', 'ALTO', 'CRITICO']
    
    for stock in niveles_stock:
        matrix[stock] = {}
        for riesgo in niveles_riesgo:
            count = len(df_completo[(df_completo['nivel_stock'] == stock) & (df_completo['riesgo_normalizado'] == riesgo)])
            matrix[stock][riesgo] = count
    
    # Generador de filas de tabla para top ventas
    def generar_filas_top_ventas():
        filas = ""
        for _, row in tv.iterrows():
            articulo = str(obtener_valor(row, 'Artículo', ''))
            nombre = str(obtener_valor(row, 'Nombre artículo', ''))
            talla = str(obtener_valor(row, 'Talla', ''))
            color = str(obtener_valor(row, 'Color', ''))
            unidades = int(obtener_valor(row, 'Ventas (unidades)', 0))
            ingresos = formatear_numero(obtener_valor(row, 'Importe ventas (€)', 0))
            beneficio = formatear_numero(obtener_valor(row, 'Beneficio (importe €)', 0))
            filas += f'''            <tr>
                <td>{articulo}</td>
                <td>{nombre}</td>
                <td>{talla}</td>
                <td>{color}</td>
                <td class="text-right">{unidades}</td>
                <td class="text-right">{ingresos}€</td>
                <td class="text-right">{beneficio}€</td>
            </tr>
'''
        return filas
    
    # Generador de filas para productos con riesgo crítico
    def generar_filas_riesgo_critico():
        filas = ""
        if '% Rotación Consumido' in df_completo.columns:
            df_critico = df_completo[df_completo['riesgo_normalizado'] == 'CRITICO'].nlargest(10, '% Rotación Consumido')
            for _, row in df_critico.iterrows():
                articulo = str(obtener_valor(row, 'Artículo', ''))
                nombre = str(obtener_valor(row, 'Nombre artículo', ''))
                talla = str(obtener_valor(row, 'Talla', ''))
                stock = int(obtener_valor(row, 'Stock Final (unidades)', 0))
                ratio = int(obtener_valor(row, '% Rotación Consumido', 0))
                filas += f'''            <tr>
                <td>{articulo}</td>
                <td>{nombre}</td>
                <td>{talla}</td>
                <td class="text-right">{stock}</td>
                <td class="text-right">{ratio}%</td>
                <td class="text-right">30%</td>
            </tr>
'''
        return filas
    
    # Generador de filas para productos problemáticos
    def generar_filas_problematicos():
        filas = ""
        for _, row in tr.iterrows():
            articulo = str(obtener_valor(row, 'Artículo', ''))
            nombre = str(obtener_valor(row, 'Nombre artículo', ''))
            talla = str(obtener_valor(row, 'Talla', ''))
            stock = int(obtener_valor(row, 'Stock Final (unidades)', 0))
            ratio = int(obtener_valor(row, '% Rotación Consumido', 0))
            valor_stock = int(stock * 30)
            filas += f'''            <tr>
                <td>{articulo}</td>
                <td>{nombre}</td>
                <td>{talla}</td>
                <td class="text-right">{stock}</td>
                <td class="text-right">{ratio}%</td>
                <td class="text-right">{valor_stock}€</td>
                <td>Liquidación urgente</td>
            </tr>
'''
        return filas
    
    # Generador de filas para productos estrella
    def generar_filas_estrella():
        filas = ""
        te_display = te[['Artículo', 'Nombre artículo', 'Talla', 'Ventas (unidades)', 'Importe ventas (€)', 'Stock Final (unidades)', 'Riesgo de Merma/ inmovilizado']].head(10)
        for _, row in te_display.iterrows():
            articulo = str(obtener_valor(row, 'Artículo', ''))
            nombre = str(obtener_valor(row, 'Nombre artículo', ''))
            talla = str(obtener_valor(row, 'Talla', ''))
            unidades = int(obtener_valor(row, 'Ventas (unidades)', 0))
            ingresos = formatear_numero(obtener_valor(row, 'Importe ventas (€)', 0))
            stock = int(obtener_valor(row, 'Stock Final (unidades)', 0))
            clasificacion = str(obtener_valor(row, 'Riesgo de Merma/ inmovilizado', ''))
            accion = 'Reposición urgente' if clasificacion == 'CERO' else ('Aumentar stock' if clasificacion == 'BAJO' else 'Mantener')
            filas += f'''            <tr>
                <td>{articulo}</td>
                <td>{nombre}</td>
                <td>{talla}</td>
                <td class="text-right">{unidades}</td>
                <td class="text-right">{ingresos}€</td>
                <td class="text-right">{stock}</td>
                <td>{accion}</td>
            </tr>
'''
        return filas
    
    # Calcular valores adicionales
    unidades_vendidas = int(df_completo[df_completo['Importe ventas (€)'] > 0]['Ventas (unidades)'].sum()) if 'Ventas (unidades)' in df_completo.columns else 0
    ticket_promedio = formatear_numero(r['ventas_totales'] / max(r['articulos_con_ventas'], 1))
    capital_liberar = int(r['capital_inmovilizado'] * 0.4)
    capital_inmov_str = formatear_numero(r['capital_inmovilizado'])
    ventas_totales_str = formatear_numero(r['ventas_totales'])
    beneficio_total_str = formatear_numero(r['beneficio_total'])
    stock_final_str = r['stock_final_total']
    margen_bruto_str = r['margen_bruto']
    
    html = f'''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Informe Final - Sección {nombre_seccion_titulo} | Enero-Febrero 2025</title>
    <style>
        :root {{
            --primary: #2E7D32;
            --secondary: #1565C0;
            --danger: #D32F2F;
            --warning: #F9A825;
            --success: #388E3C;
            --text: #37474F;
            --bg: #FAFAFA;
            --white: #FFFFFF;
        }}
        
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{
            font-family: 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            line-height: 1.6;
        }}
        
        .container {{ max-width: 1200px; margin: 0 auto; padding: 20px; }}
        
        .cover {{
            background: linear-gradient(135deg, var(--primary) 0%, #1B5E20 100%);
            color: white;
            padding: 80px 40px;
            text-align: center;
            margin-bottom: 40px;
            border-radius: 0 0 20px 20px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        }}
        
        .cover h1 {{ font-size: 2.5em; margin-bottom: 20px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }}
        .cover .subtitle {{ font-size: 1.3em; opacity: 0.9; margin-bottom: 30px; }}
        .cover .meta {{ font-size: 1em; opacity: 0.8; }}
        
        section {{
            background: var(--white);
            margin-bottom: 30px;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            page-break-inside: avoid;
        }}
        
        h2 {{ color: var(--primary); font-size: 1.6em; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid var(--primary); }}
        h3 {{ color: var(--secondary); font-size: 1.3em; margin: 20px 0 15px 0; }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }}
        
        .kpi-card {{
            background: linear-gradient(135deg, #f5f5f5 0%, #e8e8e8 100%);
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            border-left: 4px solid var(--primary);
        }}
        
        .kpi-card.danger {{ border-left-color: var(--danger); }}
        .kpi-card.warning {{ border-left-color: var(--warning); }}
        .kpi-card.success {{ border-left-color: var(--success); }}
        
        .kpi-value {{ font-size: 2em; font-weight: bold; color: var(--primary); }}
        .kpi-card.danger .kpi-value {{ color: var(--danger); }}
        .kpi-card.warning .kpi-value {{ color: var(--warning); }}
        .kpi-card.success .kpi-value {{ color: var(--success); }}
        .kpi-label {{ font-size: 0.9em; color: #666; margin-top: 5px; }}
        
        .chart-container {{ margin: 30px 0; text-align: center; }}
        .chart-title {{ font-size: 1.1em; font-weight: bold; margin-bottom: 15px; }}
        
        .table-container {{ overflow-x: auto; margin: 20px 0; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; }}
        th, td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background-color: var(--primary); color: white; font-weight: 600; position: sticky; top: 0; }}
        tr:hover {{ background-color: #f5f5f5; }}
        .text-right {{ text-align: right; }}
        .text-center {{ text-align: center; }}
        
        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.8em;
            font-weight: 600;
        }}
        
        .badge-critical {{ background: #FFEBEE; color: #C62828; }}
        .badge-high {{ background: #FFF3E0; color: #EF6C00; }}
        .badge-medium {{ background: #FFF8E1; color: #F9A825; }}
        .badge-low {{ background: #E8F5E9; color: #2E7D32; }}
        
        .matrix-grid {{
            display: grid;
            grid-template-columns: 150px repeat(4, 1fr);
            gap: 3px;
            margin: 20px 0;
        }}
        
        .matrix-cell {{ padding: 15px 10px; text-align: center; border-radius: 5px; font-size: 0.85em; }}
        .matrix-header {{ background: var(--secondary); color: white; font-weight: bold; }}
        .matrix-row-header {{ background: var(--primary); color: white; font-weight: bold; }}
        .risk-critical {{ background: #FFCDD2; }}
        .risk-high {{ background: #FFE0B2; }}
        .risk-medium {{ background: #FFF9C4; }}
        .risk-low {{ background: #C8E6C9; }}
        
        .toc {{
            background: #f8f9fa;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        
        .toc h2 {{ margin-bottom: 15px; }}
        .toc ul {{ list-style: none; columns: 2; }}
        .toc li {{ padding: 8px 0; border-bottom: 1px dashed #ddd; }}
        .toc a {{ color: var(--secondary); text-decoration: none; }}
        .toc a:hover {{ text-decoration: underline; }}
        
        footer {{
            text-align: center;
            padding: 30px;
            color: #666;
            font-size: 0.9em;
            border-top: 1px solid #ddd;
            margin-top: 40px;
        }}
        
        @media print {{
            body {{ background: white; }}
            section {{ box-shadow: none; border: 1px solid #ddd; }}
            .cover {{ background: var(--primary) !important; -webkit-print-color-adjust: exact; }}
        }}
    </style>
</head>
<body>

<div class="cover">
    <h1>INFORME FINAL</h1>
    <p class="subtitle">Seccion {nombre_seccion_titulo} - Analisis de Inventario y Ventas</p>
    <p class="meta">
        <strong>Jardineria Aranjuez (Madrid)</strong><br>
        Periodo: 1 de Enero - 28 de Febrero 2025 (59 dias)<br>
        Generado: {fecha_actual}
    </p>
</div>

<div class="container">

<!-- INDICE -->
<div class="toc">
    <h2>Indice de Contenidos</h2>
    <ul>
        <li><a href="#resumen-ejecutivo">1. Resumen Ejecutivo</a></li>
        <li><a href="#analisis-abc">2. Clasificacion ABC</a></li>
        <li><a href="#analisis-ventas">3. Analisis de Ventas</a></li>
        <li><a href="#analisis-stock">4. Analisis de Stock</a></li>
        <li><a href="#matriz-stock">5. Matriz Stock vs Rotacion</a></li>
        <li><a href="#riesgo-merma">6. Riesgo de Merma</a></li>
        <li><a href="#productos-problematicos">7. Productos Problematicos</a></li>
        <li><a href="#productos-estrella">8. Productos Estrella</a></li>
        <li><a href="#capital">9. Optimizacion de Capital</a></li>
        <li><a href="#recomendaciones">10. Recomendaciones</a></li>
    </ul>
</div>

<!-- RESUMEN EJECUTIVO -->
<section id="resumen-ejecutivo">
    <h2>1. Resumen Ejecutivo</h2>
    
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-value">{total_arts}</div>
            <div class="kpi-label">Articulos Analizados</div>
        </div>
        <div class="kpi-card success">
            <div class="kpi-value">{ventas_totales_str}€</div>
            <div class="kpi-label">Ventas Totales</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{beneficio_total_str}€</div>
            <div class="kpi-label">Beneficio</div>
        </div>
        <div class="kpi-card danger">
            <div class="kpi-value">{stock_final_str}</div>
            <div class="kpi-label">Stock Final (uds.)</div>
        </div>
    </div>
    
    <h3>Metricas Principales</h3>
    <table>
        <tr><th>Metrica</th><th>Valor</th><th>Interpretacion</th></tr>
        <tr><td>Total Articulos</td><td class="text-right">{total_arts}</td><td>SKUs unicos en catalogo</td></tr>
        <tr><td>Margen Bruto Global</td><td class="text-right">{margen_bruto_str}%</td><td>Rentabilidad saludable</td></tr>
        <tr><td>Capital Inmovilizado</td><td class="text-right">{capital_inmov_str}€</td><td>Valor del inventario</td></tr>
        <tr><td>Articulos en Riesgo Critico</td><td class="text-right" style="color: #D32F2F; font-weight: bold;">{count_critico} ({round(count_critico/total_arts*100, 1)}%)</td><td>Requieren accion inmediata</td></tr>
        <tr><td>Rupturas de Stock</td><td class="text-right" style="color: #D32F2F; font-weight: bold;">{count_cero_stock}</td><td>Oportunidades perdidas</td></tr>
        <tr><td>Productos Estrella</td><td class="text-right" style="color: #2E7D32; font-weight: bold;">{count_a}</td><td>Alta rotacion, bajo stock</td></tr>
    </table>
    
    <h3>Hallazgos Clave</h3>
    <ul>
        <li><strong>El {pct_d}% del catalogo</strong> ({count_d} articulos) no genero ventas durante el periodo</li>
        <li><strong>La clasificacion ABC</strong> muestra que {count_a} articulos ({pct_a}%) generan el {pct_ventas_a}% de los ingresos</li>
        <li><strong>El {pct_elevado}% del inventario</strong> presenta un nivel de stock ELEVADO, indicando sobreabastecimiento</li>
        <li><strong>El margen bruto global</strong> del {margen_bruto_str}% refleja una gestion rentable del negocio</li>
        <li><strong>{count_cero_stock} productos</strong> estan en ruptura de stock con demanda reciente</li>
        <li>La aplicacion de descuentos progresivos puede recuperar el 40-60% del capital en riesgo</li>
    </ul>
</section>

<!-- CLASIFICACION ABC -->
<section id="analisis-abc">
    <h2>2. Clasificacion ABC (Principio de Pareto)</h2>
    
    <div class="chart-container">
        <div class="chart-title">Distribucion de Ventas por Categoria ABC</div>
        <div style="display: flex; justify-content: center; gap: 40px; flex-wrap: wrap;">
            <div style="text-align: center;">
                <div style="width: 200px; height: 200px; border-radius: 50%; background: conic-gradient(
                    #1B5E20 0deg {pct_a * 3.6}deg, 
                    #1565C0 {pct_a * 3.6}deg {(pct_a + pct_b) * 3.6}deg, 
                    #E65100 {(pct_a + pct_b) * 3.6}deg {(pct_a + pct_b + pct_c) * 3.6}deg, 
                    #C62828 {(pct_a + pct_b + pct_c) * 3.6}deg 360deg
                ); margin: 0 auto;"></div>
                <p style="margin-top: 15px; font-weight: bold;">Distribucion Articulos</p>
            </div>
            <div style="text-align: left; max-width: 400px;">
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #1B5E20; margin-right: 10px; vertical-align: middle;"></span> Categoria A: {count_a} articulos ({pct_a}%) - Ingresos: {formatear_numero(ventas_a)}€</div>
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #1565C0; margin-right: 10px; vertical-align: middle;"></span> Categoria B: {count_b} articulos ({pct_b}%) - Ingresos: {formatear_numero(ventas_b)}€</div>
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #E65100; margin-right: 10px; vertical-align: middle;"></span> Categoria C: {count_c} articulos ({pct_c}%) - Ingresos: {formatear_numero(ventas_c)}€</div>
                <div><span style="display: inline-block; width: 20px; height: 20px; background: #C62828; margin-right: 10px; vertical-align: middle;"></span> Categoria D: {count_d} articulos ({pct_d}%) - Sin ventas</div>
            </div>
        </div>
    </div>
    
    <h3>Desglose por Categoria</h3>
    <table>
        <tr><th>Categoria</th><th>Articulos</th><th>% Articulos</th><th>Ingresos</th><th>% Ingresos</th><th>Stock Final</th><th>Acciones</th></tr>
        <tr><td><span class="badge badge-low">A - Basicos</span></td><td class="text-right">{count_a}</td><td class="text-right">{pct_a}%</td><td class="text-right">{formatear_numero(ventas_a)}€</td><td class="text-right">{pct_ventas_a}%</td><td class="text-right">{stock_a}</td><td>Mantener y optimizar</td></tr>
        <tr><td><span class="badge badge-medium">B - Complemento</span></td><td class="text-right">{count_b}</td><td class="text-right">{pct_b}%</td><td class="text-right">{formatear_numero(ventas_b)}€</td><td class="text-right">{pct_ventas_b}%</td><td class="text-right">{stock_b}</td><td>Gestion activa</td></tr>
        <tr><td><span class="badge badge-high">C - Bajo Impacto</span></td><td class="text-right">{count_c}</td><td class="text-right">{pct_c}%</td><td class="text-right">{formatear_numero(ventas_c)}€</td><td class="text-right">{pct_ventas_c}%</td><td class="text-right">{stock_c}</td><td>Evaluar continuidad</td></tr>
        <tr><td><span class="badge badge-critical">D - Sin Ventas</span></td><td class="text-right">{count_d}</td><td class="text-right">{pct_d}%</td><td class="text-right">0€</td><td class="text-right">0,0%</td><td class="text-right">{stock_d}</td><td>Liquidacion/Descatalogacion</td></tr>
        <tr style="background: #f0f0f0; font-weight: bold;">
            <td>TOTAL</td><td class="text-right">{total_arts}</td><td class="text-right">100%</td><td class="text-right">{formatear_numero(total_ventas)}€</td><td class="text-right">100%</td><td class="text-right">{stock_final_str}</td><td></td>
        </tr>
    </table>
    
    <h3>Interpretacion</h3>
    <ul>
        <li><strong>Categoria A (Basicos):</strong> {count_a} productos que representan el {pct_ventas_a}% de los ingresos. Estos son los productos estrella que deben tener prioridad en gestion de stock y reposicion.</li>
        <li><strong>Categoria B (Complemento):</strong> {count_b} productos que aportan el {pct_ventas_b}% de ingresos. Complementan la oferta y requieren gestion activa pero con menor intensidad.</li>
        <li><strong>Categoria C (Bajo Impacto):</strong> {count_c} productos con contribucion marginal del {pct_ventas_c}%. Evaluar si compensa mantenerlos en catalogo.</li>
        <li><strong>Categoria D (Sin Ventas):</strong> {count_d} productos ({pct_d}% del catalogo) sin ventas. Representan inmovilizado significativo y requieren accion inmediata de liquidacion o descatalogacion.</li>
    </ul>
</section>

<!-- ANALISIS DE VENTAS -->
<section id="analisis-ventas">
    <h2>3. Analisis de Ventas</h2>
    
    <div class="kpi-grid">
        <div class="kpi-card success">
            <div class="kpi-value">{ventas_totales_str}€</div>
            <div class="kpi-label">Ventas Totales</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{unidades_vendidas}</div>
            <div class="kpi-label">Unidades Vendidas</div>
        </div>
        <div class="kpi-card warning">
            <div class="kpi-value">{margen_bruto_str}%</div>
            <div class="kpi-label">Margen Bruto</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{ticket_promedio}€</div>
            <div class="kpi-label">Ticket Promedio</div>
        </div>
    </div>
    
    <h3>Top 15 Productos por Ingresos</h3>
    <div class="table-container">
        <table>
            <tr><th>Codigo</th><th>Nombre Articulo</th><th>Talla</th><th>Color</th><th class="text-right">Unidades</th><th class="text-right">Ingresos</th><th class="text-right">Beneficio</th></tr>
{generar_filas_top_ventas()}
        </table>
    </div>
    
    <h3>Analisis de Rentabilidad</h3>
    <table>
        <tr><th>Indicador</th><th>Valor</th><th>Evaluacion</th></tr>
        <tr><td>Ingresos Totales</td><td class="text-right">{ventas_totales_str}€</td><td>Bueno para el periodo</td></tr>
        <tr><td>Beneficio Total</td><td class="text-right">{beneficio_total_str}€</td><td>Saludable</td></tr>
        <tr><td>Margen Bruto</td><td class="text-right">{margen_bruto_str}%</td><td>Optimo (>50%)</td></tr>
        <tr><td>Unidades Vendidas</td><td class="text-right">{unidades_vendidas}</td><td>Baja rotacion</td></tr>
        <tr><td>Ticket Promedio</td><td class="text-right">{ticket_promedio}€</td><td>Moderado</td></tr>
    </table>
</section>

<!-- ANALISIS DE STOCK -->
<section id="analisis-stock">
    <h2>4. Analisis de Stock</h2>
    
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-value">{stock_final_str}</div>
            <div class="kpi-label">Stock Final (uds.)</div>
        </div>
        <div class="kpi-card warning">
            <div class="kpi-value">{capital_inmov_str}€</div>
            <div class="kpi-label">Capital Inmovilizado</div>
        </div>
        <div class="kpi-card danger">
            <div class="kpi-value">{count_elevado}</div>
            <div class="kpi-label">Stock Elevado ({pct_elevado}%)</div>
        </div>
        <div class="kpi-card success">
            <div class="kpi-value">{count_normal}</div>
            <div class="kpi-label">Stock Normal ({pct_normal}%)</div>
        </div>
    </div>
    
    <h3>Distribucion por Nivel de Stock</h3>
    <table>
        <tr><th>Nivel Stock</th><th>Articulos</th><th>% Total</th><th>Stock Final</th><th>Clasificacion</th></tr>
        <tr><td><span class="badge badge-critical">ELEVADO</span></td><td class="text-right">{count_elevado}</td><td class="text-right">{pct_elevado}%</td><td class="text-right">{stock_elevado_sum}</td><td>Sobrestock</td></tr>
        <tr><td><span class="badge badge-low">NORMAL</span></td><td class="text-right">{count_normal}</td><td class="text-right">{pct_normal}%</td><td class="text-right">{stock_normal_sum}</td><td>Optimo</td></tr>
        <tr><td><span class="badge badge-high">BAJO</span></td><td class="text-right">{count_bajo_stock}</td><td class="text-right">{pct_bajo_stock}%</td><td class="text-right">{stock_bajo_sum}</td><td>Riesgo Ruptura</td></tr>
        <tr><td><span class="badge badge-critical">CERO</span></td><td class="text-right">{count_cero_stock}</td><td class="text-right">{pct_cero_stock}%</td><td class="text-right">0</td><td>Ruptura Stock</td></tr>
    </table>
    
    <h3>Antiguedad del Stock</h3>
    <div class="chart-container">
        <div class="chart-title">Distribucion por Antiguedad de Stock</div>
        <div style="display: flex; justify-content: center; gap: 40px; flex-wrap: wrap;">
            <div style="text-align: center;">
                <div style="width: 200px; height: 200px; border-radius: 50%; background: conic-gradient({chart_gradient}); margin: 0 auto;"></div>
                <p style="margin-top: 15px; font-weight: bold;">Distribucion Antiguedad</p>
            </div>
            <div style="text-align: left; max-width: 400px;">
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #C8E6C9; margin-right: 10px; vertical-align: middle;"></span> ELEVADO: {count_elevado} articulos ({pct_elevado}%)</div>
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #FFF9C4; margin-right: 10px; vertical-align: middle;"></span> NORMAL: {count_normal} articulos ({pct_normal}%)</div>
                <div style="margin-bottom: 10px;"><span style="display: inline-block; width: 20px; height: 20px; background: #FFE0B2; margin-right: 10px; vertical-align: middle;"></span> BAJO: {count_bajo_stock} articulos ({pct_bajo_stock}%)</div>
                <div><span style="display: inline-block; width: 20px; height: 20px; background: #FFCDD2; margin-right: 10px; vertical-align: middle;"></span> CERO: {count_cero_stock} articulos ({pct_cero_stock}%)</div>
            </div>
        </div>
    </div>
</section>

<!-- MATRIZ STOCK VS ROTACION -->
<section id="matriz-stock">
    <h2>5. Matriz Stock vs Rotacion Consumida</h2>
    
    <div class="matrix-grid" style="grid-template-columns: 150px repeat(4, 1fr);">
        <div class="matrix-cell matrix-header">Stock \\ % Rotacion</div>
        <div class="matrix-cell matrix-header">BAJO<br>(<65%)</div>
        <div class="matrix-cell matrix-header">MEDIO<br>(65-100%)</div>
        <div class="matrix-cell matrix-header">ALTO<br>(100-150%)</div>
        <div class="matrix-cell matrix-header">CRITICO<br>(>150%)</div>
        
        <div class="matrix-cell matrix-row-header">ELEVADO</div>
        <div class="matrix-cell risk-low">{matrix['ELEVADO']['BAJO']}<br>articulos</div>
        <div class="matrix-cell risk-medium">{matrix['ELEVADO']['MEDIO']}<br>articulos</div>
        <div class="matrix-cell risk-high">{matrix['ELEVADO']['ALTO']}<br>articulos</div>
        <div class="matrix-cell risk-critical">{matrix['ELEVADO']['CRITICO']}<br>articulos</div>
        
        <div class="matrix-cell matrix-row-header">NORMAL</div>
        <div class="matrix-cell risk-low">{matrix['NORMAL']['BAJO']}<br>articulos</div>
        <div class="matrix-cell risk-medium">{matrix['NORMAL']['MEDIO']}<br>articulos</div>
        <div class="matrix-cell risk-high">{matrix['NORMAL']['ALTO']}<br>articulos</div>
        <div class="matrix-cell risk-critical">{matrix['NORMAL']['CRITICO']}<br>articulos</div>
        
        <div class="matrix-cell matrix-row-header">BAJO</div>
        <div class="matrix-cell risk-low">{matrix['BAJO']['BAJO']}<br>articulos</div>
        <div class="matrix-cell risk-medium">{matrix['BAJO']['MEDIO']}<br>articulos</div>
        <div class="matrix-cell risk-high">{matrix['BAJO']['ALTO']}<br>articulos</div>
        <div class="matrix-cell risk-critical">{matrix['BAJO']['CRITICO']}<br>articulos</div>
        
        <div class="matrix-cell matrix-row-header">CERO</div>
        <div class="matrix-cell risk-low">{matrix['CERO']['BAJO']}<br>articulos</div>
        <div class="matrix-cell risk-medium">{matrix['CERO']['MEDIO']}<br>articulos</div>
        <div class="matrix-cell risk-high">{matrix['CERO']['ALTO']}<br>articulos</div>
        <div class="matrix-cell risk-critical">{matrix['CERO']['CRITICO']}<br>articulos</div>
    </div>
    
    <h3>Analisis de la Matriz</h3>
    <table>
        <tr><th>Cuadrante</th><th>Articulos</th><th>Situacion</th><th>Accion Recomendada</th></tr>
        <tr><td><span class="badge badge-low">Stock ELEVADO + Riesgo BAJO</span></td><td class="text-right">{matrix['ELEVADO']['BAJO']}</td><td>Producto fresco con alta demanda</td><td>Mantener estrategia actual</td></tr>
        <tr><td><span class="badge badge-medium">Stock ELEVADO + Riesgo MEDIO</span></td><td class="text-right">{matrix['ELEVADO']['MEDIO']}</td><td>Stock abundante aproximandose a limite</td><td>Descuento preventivo 10%</td></tr>
        <tr><td><span class="badge badge-high">Stock ELEVADO + Riesgo ALTO</span></td><td class="text-right">{matrix['ELEVADO']['ALTO']}</td><td>Sobrestock con rotacion lenta</td><td>Descuento agresivo 20%</td></tr>
        <tr><td><span class="badge badge-critical">Stock ELEVADO + Riesgo CRITICO</span></td><td class="text-right">{matrix['ELEVADO']['CRITICO']}</td><td>Sobrestock critico, riesgo merma</td><td>Liquidacion urgente 30%</td></tr>
    </table>
</section>

<!-- RIESGO DE MERMA -->
<section id="riesgo-merma">
    <h2>6. Riesgo de Merma</h2>
    
    <div class="kpi-grid">
        <div class="kpi-card danger">
            <div class="kpi-value">{count_critico}</div>
            <div class="kpi-label">Riesgo Critico</div>
        </div>
        <div class="kpi-card warning">
            <div class="kpi-value">{count_alto}</div>
            <div class="kpi-label">Riesgo Alto</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{count_medio}</div>
            <div class="kpi-label">Riesgo Medio</div>
        </div>
        <div class="kpi-card success">
            <div class="kpi-value">{count_bajo_riesgo}</div>
            <div class="kpi-label">Riesgo Bajo</div>
        </div>
    </div>
    
    <h3>Articulos con Riesgo Critico de Merma</h3>
    <p>Estos articulos han superado significativamente su periodo optimo de rotacion y requieren accion inmediata:</p>
    
    <div class="table-container">
        <table>
            <tr><th>Codigo</th><th>Nombre Articulo</th><th>Talla</th><th>Stock</th><th>% Rotacion</th><th>Descuento</th></tr>
{generar_filas_riesgo_critico()}
        </table>
    </div>
    
    <h3>Plan de Pricing Dinamico</h3>
    <table>
        <tr><th>Tramo</th><th>% Rotacion Consumida</th><th>Descuento</th><th>Articulos</th><th>Accion</th></tr>
        <tr><td><span class="badge badge-low">1 - PRECIO NORMAL</span></td><td>0% - 65%</td><td>0%</td><td>{count_bajo_riesgo}</td><td>Venta a precio completo</td></tr>
        <tr><td><span class="badge badge-medium">2 - DESCUENTO PREVENTIVO</span></td><td>65% - 100%</td><td>10%</td><td>{count_medio}</td><td>Acelerar venta</td></tr>
        <tr><td><span class="badge badge-high">3 - DESCUENTO AGRESIVO</span></td><td>100% - 150%</td><td>20%</td><td>{count_alto}</td><td>Liquidacion urgente</td></tr>
        <tr><td><span class="badge badge-critical">4 - LIQUIDACION</span></td><td>> 150%</td><td>30%</td><td>{count_critico}</td><td>Recuperar valor residual</td></tr>
    </table>
    
    <h3>Recuperacion Potencial de Capital</h3>
    <ul>
        <li><strong>Con descuentos del 10%:</strong> Potencial recuperacion de {count_medio} articulos en riesgo medio</li>
        <li><strong>Con descuentos del 20%:</strong> Potencial recuperacion de {count_alto} articulos en riesgo alto</li>
        <li><strong>Con descuentos del 30%:</strong> Potencial recuperacion de {count_critico} articulos en riesgo critico</li>
        <li><strong>Total recuperable:</strong> {count_critico + count_alto + count_medio} articulos mediante estrategia de pricing dinamico</li>
    </ul>
</section>

<!-- PRODUCTOS PROBLEMATICOS -->
<section id="productos-problematicos">
    <h2>7. Productos Problematicos</h2>
    
    <p>Identificacion de articulos que requieren atencion inmediata por bajo rendimiento o alto riesgo:</p>
    
    <h3>TOP 10 Productos con Mayor Riesgo</h3>
    <div class="table-container">
        <table>
            <tr><th>Codigo</th><th>Nombre Articulo</th><th>Talla</th><th>Stock</th><th>% Rotacion</th><th>Valor Stock</th><th>Accion Sugerida</th></tr>
{generar_filas_problematicos()}
        </table>
    </div>
    
    <h3>Causas de Problematicas Identificadas</h3>
    <ul>
        <li><strong>{pct_d}% Categoria D:</strong> {count_d} productos sin ninguna venta - posible descatalogacion</li>
        <li><strong>{pct_elevado}% Stock Elevado + Riesgo Alto:</strong> {matrix['ELEVADO']['ALTO']} productos con sobreabastecimiento y baja rotacion</li>
        <li><strong>{round(count_critico/total_arts*100, 1)}% Riesgo Critico:</strong> {count_critico} productos con merma inminente</li>
        <li><strong>{pct_cero_stock}% Ruptura de Stock:</strong> {count_cero_stock} productos agotados con demanda</li>
    </ul>
</section>

<!-- PRODUCTOS ESTRELLA -->
<section id="productos-estrella">
    <h2>8. Productos Estrella</h2>
    
    <p>Articulos con alto rendimiento y gestion optima que deben mantenerse y potenciarse:</p>
    
    <h3>TOP 10 Productos con Mejor Rendimiento</h3>
    <div class="table-container">
        <table>
            <tr><th>Codigo</th><th>Nombre Articulo</th><th>Talla</th><th>Unidades Vendidas</th><th>Ingresos</th><th>Stock</th><th>Accion</th></tr>
{generar_filas_estrella()}
        </table>
    </div>
    
    <h3>Caracteristicas de Productos Estrella</h3>
    <ul>
        <li><strong>Tasa de venta del 100%:</strong> Estos productos se venden completamente durante el periodo</li>
        <li><strong>Margen bruto superior al 54%:</strong> Contribucion positiva a la rentabilidad</li>
        <li><strong>Demanda constante:</strong> Rotacion estable que permite planificar</li>
        <li><strong>Riesgo de ruptura:</strong> Muchos tienen stock bajo o cero, indicando oportunidad de aumentar compras</li>
    </ul>
    
    <h3>Recomendaciones para Productos Estrella</h3>
    <ul>
        <li><strong>Aumentar stock un 40%:</strong> Para productos con alta demanda y stock bajo</li>
        <li><strong>Mantener nivel actual:</strong> Para productos con stock equilibrado</li>
        <li><strong>Prioridad en reposicion:</strong> Estos productos deben ser los primeros en el pedido de reposicion</li>
        <li><strong>Promocion en punto de venta:</strong> Destacar estos productos para aumentar visibilidad</li>
    </ul>
</section>

<!-- OPTIMIZACION DE CAPITAL -->
<section id="capital">
    <h2>9. Optimizacion de Capital</h2>
    
    <div class="kpi-grid">
        <div class="kpi-card warning">
            <div class="kpi-value">{capital_inmov_str}€</div>
            <div class="kpi-label">Capital Total Inmovilizado</div>
        </div>
        <div class="kpi-card danger">
            <div class="kpi-value">{capital_liberar}€</div>
            <div class="kpi-label">Capital a Liberar</div>
        </div>
        <div class="kpi-card success">
            <div class="kpi-value">{count_a}</div>
            <div class="kpi-label">Prod. para Inversion</div>
        </div>
    </div>
    
    <h3>Plan de Reasignacion de Capital</h3>
    <table>
        <tr><th>Prioridad</th><th>Accion</th><th>Articulos</th><th>Capital</th><th>Impacto</th></tr>
        <tr><td class="text-center">1</td><td>Liquidacion productos Categoria D</td><td class="text-center">{count_d}</td><td class="text-right">~12000€</td><td>Recuperacion 40-60% mediante descuentos</td></tr>
        <tr><td class="text-center">2</td><td>Reducir stock Categoria C</td><td class="text-center">{count_c}</td><td class="text-right">~3500€</td><td>Eliminar productos de baja rotacion</td></tr>
        <tr><td class="text-center">3</td><td>Reposicion rupturas stock</td><td class="text-center">{count_cero_stock}</td><td>Variable</td><td>Recuperacion ventas perdidas</td></tr>
        <tr><td class="text-center">4</td><td>Aumento stock estrellas</td><td class="text-center">{count_a}</td><td>Segun demanda</td><td>+20% ventas potenciales</td></tr>
    </table>
    
    <h3>Resumen de Impacto Financiero</h3>
    <ul>
        <li><strong>Capital liberable:</strong> ~{capital_liberar}€ mediante liquidacion de productos sin ventas</li>
        <li><strong>Inversion propuesta:</strong> Reasignar capital a productos Categoria A y B</li>
        <li><strong>ROI esperado:</strong> Mejora del 15-25% en rotacion de inventario</li>
        <li><strong>Periodo recuperacion:</strong> 2-3 meses con estrategia de pricing dinamico</li>
    </ul>
</section>

<!-- RECOMENDACIONES -->
<section id="recomendaciones">
    <h2>10. Recomendaciones y Plan de Accion</h2>
    
    <h3>PRIORIDAD 1 - Acciones Inmediatas (Semana 1-2)</h3>
    <ul>
        <li>Aplicar descuento del 20-30% a {count_critico} productos con riesgo critico de merma</li>
        <li>Reposicion inmediata de {count_cero_stock} productos en ruptura de stock</li>
        <li>Implementar estrategia de pricing dinamico para {count_alto} productos en riesgo alto</li>
        <li>Revision de los {count_d} productos Categoria D sin ventas</li>
    </ul>
    
    <h3>PRIORIDAD 2 - Acciones Preventivas (Semana 3-4)</h3>
    <ul>
        <li>Implementar descuentos del 10% a {count_medio} productos en riesgo medio</li>
        <li>Aumentar compras 40% para productos estrella con bajo stock</li>
        <li>Revisar estrategia para productos Categoria C</li>
        <li>Optimizar niveles de stock segun recomendaciones por familia</li>
    </ul>
    
    <h3>PRIORIDAD 3 - Optimizacion (Mes 2)</h3>
    <ul>
        <li>Evaluar continuidad de productos no viables del catalogo</li>
        <li>Ajustar niveles de stock segun recomendaciones por familia</li>
        <li>Implementar sistema de monitoreo semanal</li>
        <li>Renegociar con proveedores para productos Categoria A</li>
    </ul>
    
    <h3>KPIs para Monitoreo</h3>
    <table>
        <tr><th>Indicador</th><th>Objetivo</th><th>Actual</th><th>Meta</th></tr>
        <tr><td>Tasa de venta semanal</td><td>>5%</td><td>Variable</td><td>Medir semanalmente</td></tr>
        <tr><td>Rotacion inventario</td><td><45 dias</td><td>Por familia</td><td>Mejorar 20%</td></tr>
        <tr><td>Productos riesgo critico</td><td><10%</td><td>{round(count_critico/total_arts*100, 1)}%</td><td>Reducir a <5%</td></tr>
        <tr><td>Rupturas de stock</td><td><5</td><td>{count_cero_stock}</td><td>Cero rupturas</td></tr>
    </table>
</section>

</div>

<footer>
    <p><strong>Informe Final - Seccion {nombre_seccion_titulo}</strong></p>
    <p>Jardineria Aranjuez (Madrid) | Periodo: Enero - Febrero 2025</p>
    <p>Generado mediante analisis automatizado de datos de inventario</p>
</footer>

</body>
</html>'''
    
    return html

def procesar_seccion(ruta_archivo, nombre_seccion):
    """Procesa un archivo de clasificación ABC+D y genera el informe HTML correspondiente."""
    print(f"\n    Procesando sección: {nombre_seccion}")
    print(f"    Archivo: {ruta_archivo}")
    
    try:
        # Leer datos del Excel
        print("    [1/4] Leyendo datos del archivo de clasificación...")
        hojas = leer_datos_clasificacion(ruta_archivo)
        
        # Combinar todas las categorías en un solo DataFrame
        print("    [2/4] Combinando datos de categorías...")
        df_completo = pd.concat(hojas.values(), ignore_index=True)
        print(f"      Total artículos: {len(df_completo)}")
        
        # Verificar columnas necesarias
        columnas_necesarias = ['Artículo', 'Nombre artículo', 'Talla', 'Color', 
                              'Importe ventas (€)', 'Beneficio (importe €)', 
                              'Stock Final (unidades)', '% Rotación Consumido',
                              'Riesgo de Merma/ inmovilizado']
        
        for col in columnas_necesarias:
            if col not in df_completo.columns:
                print(f"      Advertencia: Columna '{col}' no encontrada, se usará valor por defecto")
        
        # Calcular métricas por categoría ABC
        print("    [3/4] Calculando métricas...")
        
        # Determinar categoría ABC de cada fila según la hoja de origen
        df_a = hojas.get('CATEGORIA A – BASICOS', pd.DataFrame())
        df_b = hojas.get('CATEGORIA B – COMPLEMENTO', pd.DataFrame())
        df_c = hojas.get('CATEGORIA C – BAJO IMPACTO', pd.DataFrame())
        df_d = hojas.get('CATEGORIA D – SIN VENTAS', pd.DataFrame())
        
        # Agregar columna de categoría si no existe
        if 'categoria_abc' not in df_completo.columns:
            df_a['categoria_abc'] = 'A'
            df_b['categoria_abc'] = 'B'
            df_c['categoria_abc'] = 'C'
            df_d['categoria_abc'] = 'D'
            df_completo = pd.concat([df_a, df_b, df_c, df_d], ignore_index=True)
        
        # Calcular distribución por categoría
        dist_cat = df_completo.groupby('categoria_abc').agg({
            'Artículo': 'count', 
            'Importe ventas (€)': 'sum', 
            'Beneficio (importe €)': 'sum',
            'Stock Final (unidades)': 'sum'
        }).reset_index()
        dist_cat.columns = ['categoria', 'articulos', 'ventas', 'beneficio', 'stock']
        
        # Calcular distribución por nivel de stock (basado en Stock Final)
        if 'Stock Final (unidades)' in df_completo.columns:
            df_completo['nivel_stock'] = df_completo['Stock Final (unidades)'].apply(calcular_nivel_stock)
            dist_stock = df_completo['nivel_stock'].value_counts().reset_index()
            dist_stock.columns = ['nivel', 'articulos']
        else:
            dist_stock = pd.DataFrame({'nivel': ['NORMAL'], 'articulos': [len(df_completo)]})
        
        # Calcular distribución por nivel de riesgo (basado en % Rotación Consumido)
        if '% Rotación Consumido' in df_completo.columns:
            df_completo['riesgo_normalizado'] = df_completo['% Rotación Consumido'].apply(calcular_nivel_riesgo)
            dist_riesgo = df_completo['riesgo_normalizado'].value_counts().reset_index()
            dist_riesgo.columns = ['nivel', 'articulos']
        elif 'Riesgo de Merma/ inmovilizado' in df_completo.columns:
            df_completo['riesgo_normalizado'] = df_completo['Riesgo de Merma/ inmovilizado'].apply(normalizar_riesgo)
            dist_riesgo = df_completo['riesgo_normalizado'].value_counts().reset_index()
            dist_riesgo.columns = ['nivel', 'articulos']
        else:
            dist_riesgo = pd.DataFrame({'nivel': ['BAJO'], 'articulos': [len(df_completo)]})
            df_completo['riesgo_normalizado'] = 'BAJO'
        
        # Calcular métricas de resumen
        ventas_totales = df_completo['Importe ventas (€)'].sum()
        beneficio_total = df_completo['Beneficio (importe €)'].sum()
        stock_final_total = df_completo['Stock Final (unidades)'].sum()
        articulos_con_ventas = len(df_completo[df_completo['Importe ventas (€)'] > 0])
        articulos_sin_ventas = len(df_completo[df_completo['Importe ventas (€)'] == 0])
        
        margen_bruto = round(beneficio_total / ventas_totales * 100, 1) if ventas_totales > 0 else 0
        
        # Usar el capital inmovilizado real del archivo SPA_stock_{PERIODO}.xlsx
        # La función leer_capital_inmovilizado_stock busca primero SPA_stock_{PERIODO_FILENAME}.xlsx
        # y si no existe, usa stock.xlsx como fallback
        capital_inmovilizado_real = leer_capital_inmovilizado_stock(df_completo)
        if capital_inmovilizado_real is not None:
            capital_inmovilizado = round(capital_inmovilizado_real, 2)
        else:
            # Si no se puede leer del stock, usar el estimado como fallback
            capital_inmovilizado = round(ventas_totales * 2.5, 0)
            print(f"    ⚠ Usando valor estimado de capital inmovilizado: {capital_inmovilizado:,.2f}€")
        
        datos = {
            'resumen': {
                'total_articulos': len(df_completo),
                'articulos_con_ventas': articulos_con_ventas,
                'articulos_sin_ventas': articulos_sin_ventas,
                'ventas_totales': round(ventas_totales, 2),
                'beneficio_total': round(beneficio_total, 2),
                'stock_final_total': int(stock_final_total),
                'margen_bruto': margen_bruto,
                'capital_inmovilizado': capital_inmovilizado
            },
            'dist_categoria': dist_cat,
            'dist_stock': dist_stock,
            'dist_riesgo': dist_riesgo,
            'top_ventas': df_completo.nlargest(15, 'Importe ventas (€)')[
                ['Artículo', 'Nombre artículo', 'Talla', 'Color', 'Ventas (unidades)', 'Importe ventas (€)', 'Beneficio (importe €)']
            ] if 'Ventas (unidades)' in df_completo.columns else df_completo.nlargest(15, 'Importe ventas (€)'),
            'top_riesgo': df_completo.nlargest(10, '% Rotación Consumido')[
                ['Artículo', 'Nombre artículo', 'Talla', 'Stock Final (unidades)', '% Rotación Consumido', 'Riesgo de Merma/ inmovilizado']
            ],
            'top_estrella': df_completo[(df_completo['Importe ventas (€)'] > 0)].nlargest(15, 'Importe ventas (€)')
        }
        
        # Generar HTML
        print("    [4/4] Generando informe HTML...")
        html_informe = generar_html_informe(datos, df_completo, nombre_seccion)
        
        # Guardar archivo HTML
        nombre_salida = f"data/output/INFORME_FINAL_{nombre_seccion}_{PERIODO_FILENAME}.html"
        
        with open(nombre_salida, 'w', encoding='utf-8') as f:
            f.write(html_informe)
        
        print(f"    ✓ INFORME GENERADO: {nombre_salida}")
        return True
        
    except Exception as e:
        print(f"    ERROR procesando {nombre_seccion}: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal."""
    print("=" * 70)
    print("GENERADOR DE INFORMES ABC+D POR SECCIÓN")
    print("Vivero Aranjuez")
    print("=" * 70)
    
    # Buscar archivos de clasificación
    print("\n[1/2] Buscando archivos de clasificacion ABC+D...")
    archivos = obtener_archivos_clasificacion()
    
    if not archivos:
        print("    ERROR: No se encontraron archivos CLASIFICACION_ABC+D_*.xlsx")
        print("    Asegurate de que el script clasificacionABC.py ha generado los archivos.")
        return
    
    print(f"    Se encontraron {len(archivos)} archivo(s):")
    for archivo in archivos:
        print(f"      - {archivo}")
    
    # Procesar cada archivo
    print("\n[2/2] Procesando secciones...")
    informes_generados = 0
    errores = 0
    
    for archivo in archivos:
        nombre_seccion = extraer_nombre_seccion(archivo)
        if nombre_seccion:
            exito = procesar_seccion(archivo, nombre_seccion)
            if exito:
                informes_generados += 1
            else:
                errores += 1
        else:
            print(f"    ERROR: No se pudo extraer el nombre de sección de {archivo}")
            errores += 1
    
    # Resumen final
    print("\n" + "=" * 70)
    print("RESUMEN DE GENERACIÓN DE INFORMES")
    print("=" * 70)
    print(f"  Archivos encontrados: {len(archivos)}")
    print(f"  Informes generados: {informes_generados}")
    print(f"  Errores: {errores}")
    print("=" * 70)
    
    if informes_generados > 0:
        print("\nPROCESO COMPLETADO EXITOSAMENTE")
        print(f"Se han generado {informes_generados} informe(s) HTML:")
        
        # Recopilar todos los archivos de informe generados
        archivos_informes = []
        for archivo in archivos:
            nombre_seccion = extraer_nombre_seccion(archivo)
            if nombre_seccion:
                informe_html = f"data/output/INFORME_FINAL_{nombre_seccion}_{PERIODO_FILENAME}.html"
                archivos_informes.append(informe_html)
                print(f"  - {informe_html}")
        
        # Enviar email a Ivan con todos los informes adjuntos
        print("\nEnviando email a Ivan con los informes...")
        email_enviado = enviar_email_informes(archivos_informes)
        
        if email_enviado:
            print("  ✓ Email enviado correctamente a Ivan")
        else:
            print("  ✗ No se pudo enviar el email a Ivan")
    else:
        print("\nNo se generaron informes. Revisa los errores anteriores.")

if __name__ == "__main__":
    main()
