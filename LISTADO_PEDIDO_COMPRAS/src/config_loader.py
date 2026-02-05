#!/usr/bin/env python3
"""
Módulo de carga de configuración común para todos los scripts.
Proporciona funciones para leer la configuración desde config/config_comun.json

Este archivo es la ÚNICA fuente de verdad para todas las configuraciones.
Todos los scripts (INFORME.py, PRESENTACION.py, clasificacionABC.py)
deben importar de este módulo en lugar de tener variables hardcoded.
"""

import json
import os
from datetime import datetime
from pathlib import Path
import pandas as pd


def cargar_configuracion(ruta_config="config/config_comun.json"):
    """
    Carga la configuración desde el archivo JSON común.
    
    Args:
        ruta_config: Ruta al archivo de configuración JSON
    
    Returns:
        dict: Diccionario con todas las configuraciones cargadas
    """
    config = {}
    
    try:
        if os.path.exists(ruta_config):
            with open(ruta_config, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"  ✓ Configuración cargada desde: {ruta_config}")
        else:
            print(f"  ⚠ Archivo de configuración no encontrado: {ruta_config}")
            print("  ⚠ Usando valores por defecto")
    except Exception as e:
        print(f"  ⚠ Error al cargar configuración: {e}")
        print("  ⚠ Usando valores por defecto")
    
    return config


def calcular_periodo_desde_dataframe(df_ventas):
    """
    Calcula automáticamente el período de análisis desde un DataFrame de ventas.
    
    Args:
        df_ventas: DataFrame de ventas con columna 'Fecha'
    
    Returns:
        dict: Diccionario con FECHA_INICIO, FECHA_FIN, DIAS_PERIODO y formatos de texto
    """
    # Convertir columna Fecha a datetime si no lo es
    if df_ventas['Fecha'].dtype != 'datetime64[ns]':
        df_ventas['Fecha'] = pd.to_datetime(df_ventas['Fecha'], errors='coerce')
    
    # Calcular fechas mínima y máxima
    fecha_min = df_ventas['Fecha'].min()
    fecha_max = df_ventas['Fecha'].max()
    
    # Calcular días del período
    dias = (fecha_max - fecha_min).days + 1
    
    # Generar formatos de texto
    PERIODO_FILENAME = f"{fecha_min.strftime('%Y%m%d')}-{fecha_max.strftime('%Y%m%d')}"
    PERIODO_TEXTO = f"{fecha_min.strftime('%d de %B')} - {fecha_max.strftime('%d de %B de %Y')}"
    PERIODO_CORTO = f"{fecha_min.strftime('%B')} - {fecha_max.strftime('%B de %Y')}"
    PERIODO_EMAIL = f"{fecha_min.strftime('%d/%m/%Y')} - {fecha_max.strftime('%d/%m/%Y')}"
    
    return {
        'FECHA_INICIO': fecha_min,
        'FECHA_FIN': fecha_max,
        'DIAS_PERIODO': dias,
        'PERIODO_FILENAME': PERIODO_FILENAME,
        'PERIODO_TEXTO': PERIODO_TEXTO,
        'PERIODO_CORTO': PERIODO_CORTO,
        'PERIODO_EMAIL': PERIODO_EMAIL
    }


def obtener_configuracion_email(config=None):
    """
    Obtiene la configuración de email (compartida por todos los scripts).
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con DESTINATARIO_IVAN y SMTP_CONFIG
    """
    if config is None:
        config = cargar_configuracion()
    
    email_config = config.get('configuracion_email', {})
    
    DESTINATARIO_IVAN = email_config.get('destinatario_ivan', {
        'nombre': 'Ivan',
        'email': 'ivan.delgado@viveverde.es'
    })
    
    SMTP_CONFIG = email_config.get('smtp_config', {
        'servidor': 'smtp.serviciodecorreo.es',
        'puerto': 465,
        'remitente_email': 'ivan.delgado@viveverde.es',
        'remitente_nombre': 'Sistema de Pedidos automáticos VIVEVERDE'
    })
    
    return {
        'DESTINATARIO_IVAN': DESTINATARIO_IVAN,
        'SMTP_CONFIG': SMTP_CONFIG
    }


def obtener_configuracion_periodo_informe(config=None):
    """
    Obtiene la configuración del período para INFORME.py y PRESENTACION.py.
    NOTE: El período ahora se calcula automáticamente desde Ventas.xlsx.
    Esta función devuelve valores por defecto que serán sobrescritos por calcular_periodo_desde_dataframe().
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con FECHA_INICIO, FECHA_FIN, DIAS_PERIODO, PERIODO_FILENAME, etc.
    """
    if config is None:
        config = cargar_configuracion()
    
    periodo_config = config.get('configuracion_periodo_informe', {})
    
    # Los campos fecha_inicio y fecha_fin ya no se usan, se calculan automáticamente
    # Se mantienen para compatibilidad hacia atrás
    fecha_inicio_str = periodo_config.get('fecha_inicio', None)
    fecha_fin_str = periodo_config.get('fecha_fin', None)
    
    if fecha_inicio_str and fecha_fin_str:
        # Usar valores del config si existen (para compatibilidad)
        FECHA_INICIO = datetime.strptime(fecha_inicio_str, '%Y-%m-%d')
        FECHA_FIN = datetime.strptime(fecha_fin_str, '%Y-%m-%d')
        DIAS_PERIODO = (FECHA_FIN - FECHA_INICIO).days + 1
    else:
        # Valores por defecto que serán sobrescritos
        FECHA_INICIO = datetime(2000, 1, 1)
        FECHA_FIN = datetime(2000, 1, 1)
        DIAS_PERIODO = 1
    
    PERIODO_FILENAME = f"{FECHA_INICIO.strftime('%Y%m%d')}-{FECHA_FIN.strftime('%Y%m%d')}"
    PERIODO_TEXTO = f"{FECHA_INICIO.strftime('%d de %B')} - {FECHA_FIN.strftime('%d de %B de %Y')}"
    PERIODO_CORTO = f"{FECHA_INICIO.strftime('%B')} - {FECHA_FIN.strftime('%B de %Y')}"
    PERIODO_EMAIL = f"{FECHA_INICIO.strftime('%d/%m/%Y')} - {FECHA_FIN.strftime('%d/%m/%Y')}"
    
    return {
        'FECHA_INICIO': FECHA_INICIO,
        'FECHA_FIN': FECHA_FIN,
        'DIAS_PERIODO': DIAS_PERIODO,
        'PERIODO_FILENAME': PERIODO_FILENAME,
        'PERIODO_TEXTO': PERIODO_TEXTO,
        'PERIODO_CORTO': PERIODO_CORTO,
        'PERIODO_EMAIL': PERIODO_EMAIL
    }


def obtener_configuracion_periodo_clasificacion(config=None):
    """
    Obtiene la configuración del período para clasificacionABC.py.
    NOTE: El período ahora se calcula automáticamente desde Ventas.xlsx.
    Esta función devuelve valores por defecto que serán sobrescritos por calcular_periodo_desde_dataframe().
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con FECHA_INICIO, FECHA_FIN, DIAS_PERIODO
    """
    if config is None:
        config = cargar_configuracion()
    
    periodo_config = config.get('configuracion_periodo_clasificacion', {})
    
    # Los campos fecha_inicio y fecha_fin ya no se usan, se calculan automáticamente
    # Se mantienen para compatibilidad hacia atrás
    fecha_inicio_str = periodo_config.get('fecha_inicio', None)
    fecha_fin_str = periodo_config.get('fecha_fin', None)
    
    if fecha_inicio_str and fecha_fin_str:
        # Usar valores del config si existen (para compatibilidad)
        FECHA_INICIO = datetime.strptime(fecha_inicio_str, '%Y-%m-%d')
        FECHA_FIN = datetime.strptime(fecha_fin_str, '%Y-%m-%d')
        DIAS_PERIODO = (FECHA_FIN - FECHA_INICIO).days + 1
    else:
        # Valores por defecto que serán sobrescritos
        FECHA_INICIO = datetime(2000, 1, 1)
        FECHA_FIN = datetime(2000, 1, 1)
        DIAS_PERIODO = 1
    
    return {
        'FECHA_INICIO': FECHA_INICIO,
        'FECHA_FIN': FECHA_FIN,
        'DIAS_PERIODO': DIAS_PERIODO
    }


def obtener_configuracion_umbrales(config=None):
    """
    Obtiene la configuración de umbrales de riesgo.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con UMBRAL_RIESGO_CRITICO, UMBRAL_RIESGO_ALTO, UMBRAL_RIESGO_MEDIO
    """
    if config is None:
        config = cargar_configuracion()
    
    umbrales_config = config.get('configuracion_umbrales', {})
    
    UMBRAL_RIESGO_CRITICO = umbrales_config.get('umbral_riesgo_critico', 150)
    UMBRAL_RIESGO_ALTO = umbrales_config.get('umbral_riesgo_alto', 100)
    UMBRAL_RIESGO_MEDIO = umbrales_config.get('umbral_riesgo_medio', 65)
    
    return {
        'UMBRAL_RIESGO_CRITICO': UMBRAL_RIESGO_CRITICO,
        'UMBRAL_RIESGO_ALTO': UMBRAL_RIESGO_ALTO,
        'UMBRAL_RIESGO_MEDIO': UMBRAL_RIESGO_MEDIO
    }


def obtener_configuracion_kpis(config=None):
    """
    Obtiene la configuración de KPIs.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con KPI_OBJETIVOS y VALOR_PROMEDIO_POR_ARTICULO
    """
    if config is None:
        config = cargar_configuracion()
    
    kpis_config = config.get('configuracion_kpis', {})
    
    KPI_OBJETIVOS = kpis_config.get('kpi_objetivos', {
        'tasa_venta_semanal': '>5%',
        'rotacion_inventario_dias': '<45',
        'productos_riesgo_critico_objetivo': '<10%',
        'rupturas_stock_objetivo': '<5'
    })
    
    VALOR_PROMEDIO_POR_ARTICULO = kpis_config.get('valor_promedio_por_articulo', 50)
    
    return {
        'KPI_OBJETIVOS': KPI_OBJETIVOS,
        'VALOR_PROMEDIO_POR_ARTICULO': VALOR_PROMEDIO_POR_ARTICULO
    }


def obtener_configuracion_colores(config=None):
    """
    Obtiene la configuración de colores para Excel y reportes.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con colores de cabecera, texto y riesgo
    """
    if config is None:
        config = cargar_configuracion()
    
    colores_config = config.get('configuracion_colores', {})
    
    COLOR_CABECERA = colores_config.get('color_cabecera', '008000')
    COLOR_TEXTO_CABECERA = colores_config.get('color_texto_cabecera', 'FFFFFF')
    
    colores_riesgo_raw = colores_config.get('colores_riesgo', {})
    COLORES_RIESGO = {
        'Bajo': colores_riesgo_raw.get('Bajo', '90EE90'),
        'Medio': colores_riesgo_raw.get('Medio', 'FFFF00'),
        'Alto': colores_riesgo_raw.get('Alto', 'FFA500'),
        'Crítico': colores_riesgo_raw.get('Critico', 'FF6B6B'),
        'Cero': colores_riesgo_raw.get('Cero', '90EE90')
    }
    
    return {
        'COLOR_CABECERA': COLOR_CABECERA,
        'COLOR_TEXTO_CABECERA': COLOR_TEXTO_CABECERA,
        'COLORES_RIESGO': COLORES_RIESGO
    }


def obtener_configuracion_mascotas(config=None):
    """
    Obtiene la configuración de códigos de mascotas.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con CODIGOS_MASCOTAS_VIVO
    """
    if config is None:
        config = cargar_configuracion()
    
    mascotas_config = config.get('configuracion_mascotas', {})
    CODIGOS_MASCOTAS_VIVO = mascotas_config.get('codigos_mascotas_vivo', [
        '2104', '2204', '2305', '2405', '2504', '2606',
        '2705', '2707', '2708', '2805', '2806', '2906'
    ])
    
    return {
        'CODIGOS_MASCOTAS_VIVO': CODIGOS_MASCOTAS_VIVO
    }


def obtener_configuracion_secciones(config=None):
    """
    Obtiene la configuración de secciones.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con SECCIONES
    """
    if config is None:
        config = cargar_configuracion()
    
    secciones_config = config.get('configuracion_secciones', {})
    
    # Convertir a formato con 'descripcion' en lugar de 'label'
    SECCIONES = {}
    for nombre, datos in secciones_config.items():
        SECCIONES[nombre] = {
            'descripcion': datos.get('descripcion', nombre),
            'rangos': datos.get('rangos', [])
        }
    
    return {
        'SECCIONES': SECCIONES
    }


def obtener_configuracion_rotaciones(config=None):
    """
    Obtiene la configuración de rotaciones por familia.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con ROTACIONES_FAMILIA
    """
    if config is None:
        config = cargar_configuracion()
    
    rotaciones_raw = config.get('configuracion_rotaciones_familia', {})
    ROTACIONES_FAMILIA = {}
    
    for clave, valor in rotaciones_raw.items():
        if isinstance(valor, list) and len(valor) == 2:
            ROTACIONES_FAMILIA[clave] = (valor[0], valor[1])
        else:
            ROTACIONES_FAMILIA[clave] = ('OTROS', 90)
    
    return {
        'ROTACIONES_FAMILIA': ROTACIONES_FAMILIA
    }


def obtener_configuracion_iva(config=None):
    """
    Obtiene la configuración de IVA por familia y subfamilia.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con IVA_FAMILIA e IVA_SUBFAMILIA
    """
    if config is None:
        config = cargar_configuracion()
    
    # IVA por familia (2 dígitos)
    iva_familia_raw = config.get('configuracion_iva_familia', {})
    IVA_FAMILIA = {}
    for familia, iva in iva_familia_raw.items():
        IVA_FAMILIA[str(familia)] = int(iva)
    
    # IVA por subfamilia (4 dígitos)
    iva_subfamilia_raw = config.get('configuracion_iva_subfamilia', {})
    IVA_SUBFAMILIA = {}
    for subfamilia, iva in iva_subfamilia_raw.items():
        IVA_SUBFAMILIA[str(subfamilia)] = int(iva)
    
    return {
        'IVA_FAMILIA': IVA_FAMILIA,
        'IVA_SUBFAMILIA': IVA_SUBFAMILIA
    }


def get_abc_config(config=None):
    """
    Obtiene la configuración ABC completa para clasificacionABC.py.
    NOTE: Las fechas de período ya no se devuelven ya que se calculan automáticamente
    desde Ventas.xlsx usando calcular_periodo_desde_dataframe().
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con todas las configuraciones ABC (excepto fechas)
    """
    if config is None:
        config = cargar_configuracion()
    
    colores = obtener_configuracion_colores(config)
    mascotas = obtener_configuracion_mascotas(config)
    secciones = obtener_configuracion_secciones(config)
    rotaciones = obtener_configuracion_rotaciones(config)
    iva = obtener_configuracion_iva(config)
    
    return {
        'codigos_mascotas_vivo': mascotas['CODIGOS_MASCOTAS_VIVO'],
        'colores_riesgo': colores['COLORES_RIESGO'],
        'color_cabecera': colores['COLOR_CABECERA'],
        'color_texto_cabecera': colores['COLOR_TEXTO_CABECERA']
    }


def calcular_periodo_ventas(ventas_df):
    """
    Calcula el período de análisis desde el DataFrame de ventas.
    
    Args:
        ventas_df: DataFrame de ventas con columna 'Fecha'
    
    Returns:
        tuple: (FECHA_INICIO, FECHA_FIN, DIAS_PERIODO)
    """
    # Usar la función ya definida que retorna un diccionario
    resultado = calcular_periodo_desde_dataframe(ventas_df)
    
    return (
        resultado['FECHA_INICIO'],
        resultado['FECHA_FIN'],
        resultado['DIAS_PERIODO']
    )


# Funciones wrapper con nombres compatibles con clasificacionABC.py

def get_secciones_config(config=None):
    """Wrapper para obtener_configuracion_secciones"""
    return obtener_configuracion_secciones(config)['SECCIONES']


def get_encargados_config(config=None):
    """Wrapper para obtener_configuracion_encargados"""
    return obtener_configuracion_encargados(config)


def get_smtp_config(config=None):
    """Wrapper para obtener SMTP_CONFIG desde obtener_configuracion_email"""
    return obtener_configuracion_email(config)['SMTP_CONFIG']


def get_rotaciones_familia(config=None):
    """Wrapper para obtener_configuracion_rotaciones"""
    return obtener_configuracion_rotaciones(config)['ROTACIONES_FAMILIA']


def get_iva_familia(config=None):
    """Wrapper para obtener IVA_FAMILIA desde obtener_configuracion_iva"""
    return obtener_configuracion_iva(config)['IVA_FAMILIA']


def get_iva_subfamilia(config=None):
    """Wrapper para obtener IVA_SUBFAMILIA desde obtener_configuracion_iva"""
    return obtener_configuracion_iva(config)['IVA_SUBFAMILIA']


def obtener_configuracion_encargados(config=None):
    """
    Obtiene la configuración de encargados por sección.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con ENCARGADOS
    """
    if config is None:
        config = cargar_configuracion()
    
    encargado_config = config.get('configuracion_encargados', {})
    
    # Convertir claves de sección a minúsculas
    ENCARGADOS = {}
    for seccion, datos in encargado_config.items():
        ENCARGADOS[seccion.lower()] = {
            'nombre': datos.get('nombre', 'Ivan'),
            'email': datos.get('email', 'ivan.delgado@viveverde.es')
        }
    
    return {
        'ENCARGADOS': ENCARGADOS
    }


def obtener_configuracion_email_textos(config=None):
    """
    Obtiene los textos de emails.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario con asuntos y cuerpos de emails
    """
    if config is None:
        config = cargar_configuracion()
    
    texto_config = config.get('configuracion_texto_email', {})
    
    return {
        'ASUNTO_INFORME': texto_config.get('asunto_informe', 'VIVEVERDE: Informes de ClasificacionABC+D de cada sección del periodo {periodo}'),
        'CUERPO_INFORME': texto_config.get('cuerpo_informe', 'Buenos días {nombre},\n\nTe adjunto en este correo los informes de Clasificación ABC+D de cada sección.\n\nAtentamente,\n\nSistema de Pedidos automáticos VIVEVERDE.'),
        'ASUNTO_PRESENTACION': texto_config.get('asunto_presentacion', 'VIVEVERDE: Presentacion de ClasificacionABC+D de cada sección del periodo {periodo}'),
        'CUERPO_PRESENTACION': texto_config.get('cuerpo_presentacion', 'Buenos días {nombre},\n\nTe adjunto en este correo las presentaciones de Clasificación ABC+D de cada sección.\n\nAtentamente,\n\nSistema de Pedidos automáticos VIVEVERDE.'),
        'ASUNTO_CLASIFICACION': texto_config.get('asunto_clasificacion', 'VIVEVERDE: listado ClasificacionABC+D de {seccion} del periodo {periodo}'),
        'CUERPO_CLASIFICACION': texto_config.get('cuerpo_clasificacion', 'Buenos días {nombre},\n\nTe adjunto en este correo el listado Clasificación ABC+D de {seccion} para que lo analices y te aprendas cuales son los artículos de cada categoría:\n\n- Artículos que no te deben faltar nunca (Categoria A).\n- Artículos que confeccionan el complemento de gama (Categoría B).\n- Artículos que tienen una presencia mínima en las ventas de tu sección (Categoría C).\n- Artículos que no debemos tener en tienda (Categoria D).\n\nPon en práctica el listado.\n\nAtentamente,\n\nSistema de Pedidos automáticos VIVEVERDE.')
    }


def obtener_configuracion_completa_informe(config=None):
    """
    Obtiene toda la configuración necesaria para INFORME.py.
    NOTE: El período debe calcularse con calcular_periodo_desde_dataframe() usando el DataFrame de ventas.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario completo con todas las configuraciones de INFORME.py
    """
    if config is None:
        config = cargar_configuracion()
    
    email = obtener_configuracion_email(config)
    periodo = obtener_configuracion_periodo_informe(config)
    umbrales = obtener_configuracion_umbrales(config)
    kpis = obtener_configuracion_kpis(config)
    
    return {
        **email,
        **periodo,
        **umbrales,
        **kpis
    }


def obtener_configuracion_completa_presentacion(config=None):
    """
    Obtiene toda la configuración necesaria para PRESENTACION.py.
    NOTE: El período debe calcularse con calcular_periodo_desde_dataframe() usando el DataFrame de ventas.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario completo con todas las configuraciones de PRESENTACION.py
    """
    if config is None:
        config = cargar_configuracion()
    
    email = obtener_configuracion_email(config)
    periodo = obtener_configuracion_periodo_informe(config)
    
    return {
        **email,
        **periodo
    }


def obtener_configuracion_completa_clasificacion(config=None):
    """
    Obtiene toda la configuración necesaria para clasificacionABC.py.
    NOTE: El período debe calcularse con calcular_periodo_desde_dataframe() usando el DataFrame de ventas.
    
    Args:
        config: Diccionario de configuración (opcional)
    
    Returns:
        dict: Diccionario completo con todas las configuraciones de clasificacionABC.py
    """
    if config is None:
        config = cargar_configuracion()
    
    email = obtener_configuracion_email(config)
    periodo = obtener_configuracion_periodo_clasificacion(config)
    colores = obtener_configuracion_colores(config)
    mascotas = obtener_configuracion_mascotas(config)
    secciones = obtener_configuracion_secciones(config)
    rotaciones = obtener_configuracion_rotaciones(config)
    iva = obtener_configuracion_iva(config)
    encargados = obtener_configuracion_encargados(config)
    
    return {
        **email,
        **periodo,
        **colores,
        **mascotas,
        **secciones,
        **rotaciones,
        **iva,
        **encargados
    }


if __name__ == "__main__":
    # Test de carga de configuración
    print("=" * 60)
    print("TEST DE CARGA DE CONFIGURACIÓN")
    print("=" * 60)
    
    config = cargar_configuracion()
    
    # Test de configuración de email
    print("\n--- Configuración de Email ---")
    email = obtener_configuracion_email(config)
    print(f"Remitente: {email['SMTP_CONFIG']['remitente_email']}")
    print(f"Destinatario: {email['DESTINATARIO_IVAN']['email']}")
    
    # Test de configuración de período informe
    print("\n--- Período para INFORME.py y PRESENTACION.py (valores por defecto) ---")
    periodo_informe = obtener_configuracion_periodo_informe(config)
    print(f"Inicio: {periodo_informe['FECHA_INICIO'].strftime('%d/%m/%Y')}")
    print(f"Fin: {periodo_informe['FECHA_FIN'].strftime('%d/%m/%Y')}")
    print(f"Días: {periodo_informe['DIAS_PERIODO']}")
    print("NOTA: El período real se calculará automáticamente desde Ventas.xlsx")
    
    # Test de configuración de período clasificación
    print("\n--- Período para clasificacionABC.py (valores por defecto) ---")
    periodo_clasif = obtener_configuracion_periodo_clasificacion(config)
    print(f"Inicio: {periodo_clasif['FECHA_INICIO'].strftime('%d/%m/%Y')}")
    print(f"Fin: {periodo_clasif['FECHA_FIN'].strftime('%d/%m/%Y')}")
    print(f"Días: {periodo_clasif['DIAS_PERIODO']}")
    print("NOTA: El período real se calculará automáticamente desde Ventas.xlsx")
    
    # Test de configuración de umbrales
    print("\n--- Umbrales de Riesgo ---")
    umbrales = obtener_configuracion_umbrales(config)
    print(f"Crítico: {umbrales['UMBRAL_RIESGO_CRITICO']}%")
    print(f"Alto: {umbrales['UMBRAL_RIESGO_ALTO']}%")
    print(f"Medio: {umbrales['UMBRAL_RIESGO_MEDIO']}%")
    
    # Test de configuración de KPIs
    print("\n--- KPIs ---")
    kpis = obtener_configuracion_kpis(config)
    print(f"Objetivos: {kpis['KPI_OBJETIVOS']}")
    print(f"Valor promedio artículo: {kpis['VALOR_PROMEDIO_POR_ARTICULO']}€")
    
    # Test de secciones
    print("\n--- Secciones ---")
    secciones = obtener_configuracion_secciones(config)
    print(f"Total secciones: {len(secciones['SECCIONES'])}")
    print(f"Secciones: {list(secciones['SECCIONES'].keys())}")
    
    # Test de rotaciones
    print("\n--- Rotaciones (muestra) ---")
    rotaciones = obtener_configuracion_rotaciones(config)
    print(f"Familia 11: {rotaciones['ROTACIONES_FAMILIA'].get('11', ('Desconocida', 0))}")
    print(f"Familia 81: {rotaciones['ROTACIONES_FAMILIA'].get('81', ('Desconocida', 0))}")
    
    print("\n" + "=" * 60)
    print("✓ TODAS LAS CONFIGURACIONES CARGADAS CORRECTAMENTE")
    print("NOTA: El período se calculará automáticamente desde Ventas.xlsx")
    print("=" * 60)
