#!/usr/bin/env python3
"""
Sistema de Pedidos de Compra - Vivero Aranjuez V2

Sistema modular para la generación automática de pedidos de compra
con generación semanal programada, persistencia de estado, sistema de corrección (FASE 2)
y envío automático de emails a los responsables de cada sección.

Este es el módulo principal que coordina todos los componentes del sistema:
- DataLoader: Carga de datos de entrada
- StateManager: Persistencia de estado entre ejecuciones
- ForecastEngine: Cálculo de pedidos (FASE 1)
- CorrectionDataLoader: Carga de datos de corrección (FASE 2)
- CorrectionEngine: Motor de corrección de pedidos (FASE 2)
- OrderGenerator: Generación de archivos de salida
- SchedulerService: Control de ejecución programada
- EmailService: Envío de emails a responsables

MODO DE USO:
- Ejecución normal (programada): python main.py
- Ejecución forzada para una semana específica: python main.py --semana 15
- Ejecución en modo continuo (esperando horario): python main.py --continuo
- Mostrar información del estado: python main.py --status
- Solo corrección (FASE 2): python main.py --correccion --semana 15
- Con corrección habilitada: python main.py --semana 15 --con-correccion
- Sin envío de emails: python main.py --semana 15 --sin-email

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-03 (Actualizado con Email Service)
"""

import sys
import os
import json
import logging
import argparse
from datetime import datetime
from typing import Optional, Dict, Any, Tuple, List

# Añadir el directorio actual al path para imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Imports de los módulos del sistema
from src.data_loader import DataLoader
from src.state_manager import StateManager
from src.forecast_engine import ForecastEngine
from src.order_generator import OrderGenerator
from src.scheduler_service import SchedulerService, EstadoEjecucion

# Imports de FASE 2 - Sistema de Corrección
from src.correction_data_loader import CorrectionDataLoader
from src.correction_engine import CorrectionEngine, crear_correction_engine

# Imports de EMAIL Service
from src.email_service import EmailService, crear_email_service

# Importar pandas
import pandas as pd

# ============================================
# CONFIGURACIÓN DE LOGGING
# ============================================

def configurar_logging(nivel: int = logging.INFO, log_file: Optional[str] = None) -> logging.Logger:
    """
    Configura el sistema de logging del sistema.
    
    Args:
        nivel (int): Nivel de logging (logging.INFO, logging.DEBUG, etc.)
        log_file (Optional[str]): Ruta del archivo de log (opcional)
    
    Returns:
        logging.Logger: Instancia del logger configurada
    """
    # Crear formato
    formato = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Configurar logger raíz
    logger = logging.getLogger()
    logger.setLevel(nivel)
    
    # Limpiar handlers existentes
    logger.handlers = []
    
    # Handler de consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)
    logger.addHandler(console_handler)
    
    # Handler de archivo (si se especifica)
    if log_file:
        # Crear directorio de logs si no existe
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setFormatter(formato)
        logger.addHandler(file_handler)
    
    return logger


# ============================================
# CARGA DE CONFIGURACIÓN
# ============================================

def cargar_configuracion(ruta: str = 'config/config.json') -> Optional[Dict[str, Any]]:
    """
    Carga la configuración desde el archivo JSON.
    
    Args:
        ruta (str): Ruta al archivo de configuración
    
    Returns:
        Optional[Dict]: Configuración cargada o None si hay error
    """
    try:
        # Obtener directorio base (donde está main.py)
        dir_base = os.path.dirname(os.path.abspath(__file__))
        ruta_completa = os.path.join(dir_base, ruta)
        
        if not os.path.exists(ruta_completa):
            print(f"ERROR: No se encontró el archivo de configuración: {ruta_completa}")
            return None
        
        with open(ruta_completa, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print(f"Configuración cargada desde: {ruta}")
        return config
        
    except Exception as e:
        print(f"ERROR al cargar configuración: {str(e)}")
        return None


# ============================================
# FASE 2: FUNCIONES DE CORRECCIÓN
# ============================================

def verificar_archivos_correccion(config: Dict[str, Any], semana: int) -> Dict[str, bool]:
    """
    Verifica qué archivos de corrección están disponibles para una semana.
    
    Args:
        config (Dict): Configuración del sistema
        semana (int): Número de semana
    
    Returns:
        Dict[str, bool]: Disponibilidad de cada archivo
    """
    dir_entrada = config.get('rutas', {}).get('directorio_entrada', './data/input')
    archivos_correccion = config.get('archivos_correccion', {})
    
    disponibilidad = {
        'stock': False,
        'ventas': False,
        'compras': False
    }
    
    # Patrones de búsqueda para cada archivo
    patrones = {
        'stock': ['Stock_actual.xlsx', f'Stock_semana_{semana}.xlsx'],
        'ventas': [f'Ventas_semana_{semana}.xlsx', f'Ventas_Semana_{semana}.xlsx', 'Ventas_semana.xlsx'],
        'compras': [f'Compras_semana_{semana}.xlsx', f'Compras_Semana_{semana}.xlsx', 'Compras_semana.xlsx']
    }
    
    for tipo, patrones_archivo in patrones.items():
        for patron in patrones_archivo:
            ruta = os.path.join(dir_entrada, patron)
            if os.path.exists(ruta):
                disponibilidad[tipo] = True
                logger.info(f"Archivo de {tipo} encontrado: {patron}")
                break
    
    return disponibilidad


def aplicar_correccion_pedido(
    pedido_teorico: pd.DataFrame,
    semana: int,
    config: Dict[str, Any],
    parametros_abc: Optional[Dict[str, Any]] = None
) -> Tuple[Optional[pd.DataFrame], Dict[str, Any]]:
    """
    Aplica la corrección FASE 2 a un pedido teórico.
    
    Args:
        pedido_teorico (pd.DataFrame): Pedido generado en FASE 1
        semana (int): Número de semana
        config (Dict): Configuración del sistema
        parametros_abc (Optional[Dict]): Parámetros ABC para el motor de corrección
    
    Returns:
        Tuple[Optional[pd.DataFrame], Dict]: (Pedido corregido, Métricas de corrección)
    """
    logger.info("\n" + "=" * 60)
    logger.info("FASE 2: APLICANDO CORRECCIÓN AL PEDIDO")
    logger.info("=" * 60)
    
    # Verificar si la corrección está habilitada
    params_correccion = config.get('parametros_correccion', {})
    if not params_correccion.get('habilitar_correccion', True):
        logger.info("Corrección deshabilitada en configuración. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False}
    
    # Verificar archivos de corrección disponibles
    disponibilidad = verificar_archivos_correccion(config, semana)
    
    if not any(disponibilidad.values()):
        logger.warning("No se encontraron archivos de corrección. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_archivos'}
    
    logger.info(f"Archivos de corrección disponibles: {disponibilidad}")
    
    try:
        # Inicializar CorrectionDataLoader
        correction_loader = CorrectionDataLoader(config)
        
        # Cargar datos de corrección
        datos_correccion = correction_loader.cargar_datos_correccion(semana)
        
        # Verificar si hay datos cargados
        datos_cargados = sum(1 for v in datos_correccion.values() if v is not None)
        if datos_cargados == 0:
            logger.warning("No se pudieron cargar datos de corrección. Usando pedido teórico.")
            return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_datos'}
        
        # Fusionar datos de corrección con pedido teórico
        pedido_fusionado = correction_loader.merge_con_pedido_teorico(
            pedido_teorico, datos_correccion
        )
        
        # Inicializar CorrectionEngine
        config_abc = {
            'pesos_categoria': config.get('parametros', {}).get('pesos_categoria', {})
        }
        politica_stock = params_correccion.get('stock_minimo_por_categoria', {
            'A': 1.5, 'B': 1.0, 'C': 0.5, 'D': 0.0
        })
        
        engine = crear_correction_engine(
            config_abc=config_abc,
            politica_stock_minimo=politica_stock
        )
        
        # Aplicar corrección
        pedido_corregido = engine.aplicar_correccion_dataframe(
            pedido_fusionado,
            columna_pedido='Unidades_Pedido',
            columna_stock_minimo='Stock_Minimo_Objetivo',
            columna_stock_real='Stock_Fisico',
            columna_categoria='Categoria',
            columna_ventas_reales='Unidades_Vendidas',
            columna_ventas_objetivo='Ventas_Objetivo',
            columna_compras_reales='Unidades_Recibidas',
            columna_compras_sugeridas='Unidades_Pedido'
        )
        
        # Calcular métricas
        metricas = engine.calcular_metricas_correccion(
            pedido_corregido,
            columna_pedido_original='Unidades_Pedido',
            columna_pedido_corregido='Pedido_Corregido',
            columna_ventas_reales='Unidades_Vendidas',
            columna_ventas_objetivo='Ventas_Objetivo'
        )
        metricas['correccion_aplicada'] = True
        metricas['datos_cargados'] = datos_cargados
        
        # Generar alertas
        alertas = engine.generar_alertas(pedido_corregido)
        if alertas:
            metricas['alertas'] = alertas
            logger.warning("ALERTAS GENERADAS:")
            for alerta in alertas:
                logger.warning(f"  [{alerta['nivel']}] {alerta['mensaje']}")
        
        # Resumen de corrección
        logger.info("\nRESUMEN DE CORRECCIÓN:")
        logger.info(f"  Artículos corregidos: {metricas['articulos_corregidos']}/{metricas['total_articulos']}")
        logger.info(f"  Porcentaje corregido: {metricas['porcentaje_corregidos']:.1f}%")
        logger.info(f"  Diferencia unidades: {int(metricas['diferencia_unidades']):+d}")
        logger.info(f"  Porcentaje cambio: {metricas['porcentaje_cambio']:+.1f}%")
        
        return pedido_corregido, metricas
        
    except Exception as e:
        logger.error(f"Error al aplicar corrección: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'error', 'error': str(e)}


def generar_archivo_pedido_corregido(
    pedido_corregido: pd.DataFrame,
    semana: int,
    seccion: str,
    parametros_seccion: Dict[str, Any],
    config: Dict[str, Any],
    order_generator: OrderGenerator
) -> Optional[str]:
    """
    Genera el archivo Excel con el pedido corregido.
    
    Args:
        pedido_corregido (pd.DataFrame): Pedido tras aplicar corrección
        semana (int): Número de semana
        seccion (str): Nombre de la sección
        parametros_seccion (Dict): Parámetros de la sección
        config (Dict): Configuración del sistema
        order_generator (OrderGenerator): Generador de archivos
    
    Returns:
        Optional[str]: Ruta del archivo generado o None si hay error
    """
    try:
        # Calcular fechas
        from datetime import datetime, timedelta
        fecha_base = datetime.now()
        
        # Calcular fecha del lunes de la semana
        dia_semana = fecha_base.weekday()
        dias_hasta_lunes = (7 - dia_semana) % 7
        fecha_lunes = fecha_base + timedelta(days=dias_hasta_lunes + (7 * ((semana - fecha_base.isocalendar()[1]) % 52)))
        
        # Usar fecha actual si es futuro
        if fecha_lunes < datetime.now():
            fecha_lunes = datetime.now()
        
        fecha_lunes_str = fecha_lunes.strftime('%Y-%m-%d')
        
        # Generar nombre de archivo con sufijo "_CORREGIDO"
        dir_salida = config.get('rutas', {}).get('directorio_salida', './data/output')
        nombre_archivo = f"Pedido_Semana_{semana}_{fecha_lunes_str}_{seccion}_CORREGIDO.xlsx"
        ruta_archivo = os.path.join(dir_salida, nombre_archivo)
        
        # Preparar datos para exportación
        df_exportar = pedido_corregido.copy()
        
        # Renombrar columnas para claridad
        renombrar = {
            'Unidades_Pedido': 'Pedido_Teorico',
            'Pedido_Corregido': 'Pedido_Final',
            'Stock_Minimo_Objetivo': 'Stock_Minimo',
            'Stock_Fisico': 'Stock_Real',
            'Unidades_Vendidas': 'Ventas_Reales',
            'Unidades_Recibidas': 'Compras_Recibidas',
            'Diferencia_Stock': 'Ajuste_Stock',
            'Razon_Correccion': 'Correccion_Aplicada'
        }
        
        for col_vieja, col_nueva in renombrar.items():
            if col_vieja in df_exportar.columns:
                df_exportar.rename(columns={col_vieja: col_nueva}, inplace=True)
        
        # Ordenar columnas
        columnas_orden = [
            'Código artículo', 'Nombre artículo', 'Talla', 'Color', 'Categoria',
            'Pedido_Teorico', 'Stock_Minimo', 'Stock_Real', 'Ajuste_Stock',
            'Pedido_Final', 'Correccion_Aplicada', 'Escenario',
            'Ventas_Reales', 'Ventas_Objetivo', 'Compras_Recibidas',
            'PVP', 'Coste', 'Proveedor', 'Unidades_ABC', 'Ventas_Objetivo'
        ]
        
        columnas_finales = [col for col in columnas_orden if col in df_exportar.columns]
        df_exportar = df_exportar[columnas_finales]
        
        # Guardar archivo
        df_exportar.to_excel(ruta_archivo, index=False, sheet_name=seccion.capitalize())
        
        logger.info(f"Archivo de pedido corregido generado: {nombre_archivo}")
        return ruta_archivo
        
    except Exception as e:
        logger.error(f"Error al generar archivo corregido: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None


# ============================================
# EMAIL SERVICE: FUNCIONES DE ENVÍO DE EMAILS
# ============================================

def enviar_emails_pedidos(
    semana: int,
    config: Dict[str, Any],
    archivos_por_seccion: Dict[str, List[str]]
) -> Dict[str, Any]:
    """
    Envía los archivos de pedido por email a los responsables de cada sección.
    
    Args:
        semana (int): Número de semana procesada
        config (Dict): Configuración del sistema
        archivos_por_seccion (Dict): Mapeo sección -> lista de archivos generados
    
    Returns:
        Dict: Resultado del envío de emails
    """
    logger.info("\n" + "=" * 60)
    logger.info("ENVÍO DE EMAILS A RESPONSABLES")
    logger.info("=" * 60)
    
    # Verificar si el envío de emails está habilitado
    email_config = config.get('email', {})
    if not email_config.get('habilitar_envio', True):
        logger.info("Envío de emails deshabilitado en configuración.")
        return {'exito': False, 'razon': 'deshabilitado'}
    
    try:
        # Crear servicio de email
        email_service = crear_email_service(config)
        
        # Verificar configuración
        verificacion = email_service.verificar_configuracion()
        if not verificacion['valido']:
            logger.warning("Problemas en la configuración de email:")
            for problema in verificacion['problemas']:
                logger.warning(f"  - {problema}")
            
            # Si falta la contraseña, no continuamos
            if any('EMAIL_PASSWORD' in p for p in verificacion['problemas']):
                logger.error("No se puede enviar emails sin configurar la variable EMAIL_PASSWORD")
                return {'exito': False, 'razon': 'sin_password'}
        
        # Enviar email para cada sección
        resultados = {}
        emails_enviados = 0
        emails_fallidos = 0
        
        for seccion, archivos in archivos_por_seccion.items():
            if not archivos:
                logger.info(f"Sin archivos para la sección {seccion}. Saltando.")
                continue
            
            logger.info(f"\nEnviando email para sección: {seccion}")
            logger.info(f"Archivos: {archivos}")
            
            resultado = email_service.enviar_pedido_por_seccion(
                semana=semana,
                seccion=seccion,
                archivos=archivos
            )
            
            resultados[seccion] = resultado
            
            if resultado.get('exito', False):
                emails_enviados += 1
                logger.info(f"✓ Email enviado exitosamente a {seccion}")
            else:
                emails_fallidos += 1
                logger.error(f"✗ Error al enviar email a {seccion}: {resultado.get('error', 'Error desconocido')}")
        
        # Resumen del envío
        logger.info("\n" + "=" * 60)
        logger.info("RESUMEN DE ENVÍO DE EMAILS")
        logger.info("=" * 60)
        logger.info(f"Emails enviados exitosamente: {emails_enviados}")
        logger.info(f"Emails fallidos: {emails_fallidos}")
        logger.info(f"Secciones procesadas: {len(resultados)}")
        
        return {
            'exito': emails_enviados > 0,
            'emails_enviados': emails_enviados,
            'emails_fallidos': emails_fallidos,
            'resultados': resultados
        }
        
    except Exception as e:
        logger.error(f"Error al enviar emails: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {'exito': False, 'error': str(e)}


def agrupar_archivos_por_seccion(
    archivos_generados: List[str],
    config: Dict[str, Any]
) -> Dict[str, List[str]]:
    """
    Agrupa los archivos generados por sección para el envío de emails.
    
    Args:
        archivos_generados (List[str]): Lista de rutas de archivos generados
        config (Dict): Configuración del sistema
    
    Returns:
        Dict[str, List[str]]: Mapeo sección -> lista de archivos
    """
    archivos_por_seccion = {}
    
    # Obtener directorio de salida
    dir_salida = config.get('rutas', {}).get('directorio_salida', './data/output')
    
    for archivo in archivos_generados:
        if not archivo:
            continue
        
        # Extraer nombre del archivo
        nombre_archivo = os.path.basename(archivo)
        
        # Determinar la sección a partir del nombre del archivo
        # Formato esperado: Pedido_Semana_{semana}_{fecha}_{seccion}.xlsx
        partes = nombre_archivo.replace('.xlsx', '').split('_')
        
        if len(partes) >= 4:
            # La sección es la última parte del nombre
            seccion = partes[-1]
            
            # Filtrar solo archivos principales (no resúmenes ni corregidos si es necesario)
            if 'RESUMEN' in nombre_archivo.upper():
                continue  # Saltar resúmenes para envío individual
            
            if seccion not in archivos_por_seccion:
                archivos_por_seccion[seccion] = []
            
            archivos_por_seccion[seccion].append(archivo)
    
    logger.debug(f"[DEBUG] Archivos agrupados por sección: {archivos_por_seccion}")
    return archivos_por_seccion


# ============================================
# PROCESO PRINCIPAL DE GENERACIÓN DE PEDIDO
# ============================================

def procesar_pedido_semana(
    semana: int, 
    config: Dict[str, Any], 
    state_manager: StateManager,
    forzar: bool = False,
    aplicar_correccion: bool = True,
    enviar_email: bool = True
) -> Tuple[bool, Optional[str], int, float, Dict[str, Any], Dict[str, Any]]:
    """
    Procesa el pedido para una semana específica (FASE 1 + FASE 2 opcional).
    
    Args:
        semana (int): Número de semana a procesar
        config (Dict): Configuración del sistema
        state_manager (StateManager): Gestor de estado
        forzar (bool): Si True, fuerza el procesamiento aunque ya esté hecho
        aplicar_correccion (bool): Si True, aplica la corrección FASE 2
        enviar_email (bool): Si True, envía los archivos por email
    
    Returns:
        Tuple[bool, Optional[str], int, float, Dict, Dict]: 
            (éxito, archivo_generado, articulos, importe, metricas_correccion, resultado_email)
    """
    logger.info("=" * 70)
    logger.info(f"PROCESANDO PEDIDO PARA SEMANA {semana}")
    logger.info("=" * 70)
    
    # Mensaje sobre corrección
    if aplicar_correccion:
        logger.info("MODO: FASE 1 (Forecast) + FASE 2 (Corrección)")
    else:
        logger.info("MODO: Solo FASE 1 (Forecast) - Corrección deshabilitada")
    
    # Inicializar componentes FASE 1
    data_loader = DataLoader(config)
    forecast_engine = ForecastEngine(config)
    order_generator = OrderGenerator(config)
    scheduler = SchedulerService(config)
    
    # Calcular fechas de la semana
    fecha_lunes, fecha_domingo, fecha_archivo = scheduler.calcular_fechas_semana_pedido(semana)
    logger.info(f"Período de la semana: {fecha_lunes} al {fecha_domingo}")
    
    # Obtener stock acumulado del estado
    stock_acumulado = state_manager.obtener_stock_acumulado()
    logger.info(f"Stock acumulado cargado: {len(stock_acumulado)} artículos")
    
    # Procesar cada sección configurada
    secciones = config.get('secciones_activas', [])
    pedidos_totales = {}
    pedidos_corregidos = {}
    datos_semanales = {}
    articulos_totales = 0
    importe_total = 0.0
    metricas_correccion_total = {}
    
    archivos_generados = []
    
    for seccion in secciones:
        logger.info(f"\n{'=' * 50}")
        logger.info(f"SECCION: {seccion.upper()}")
        logger.info(f"{'=' * 50}")
        
        try:
            # Leer datos de la sección
            abc_df, ventas_df, costes_df = data_loader.leer_datos_seccion(seccion)
            
            logger.debug(f"[DEBUG] abc_df: {len(abc_df) if abc_df is not None else 0} registros")
            logger.debug(f"[DEBUG] ventas_df: {len(ventas_df) if ventas_df is not None else 0} registros")
            logger.debug(f"[DEBUG] costes_df: {len(costes_df) if costes_df is not None else 0} registros")
            
            if abc_df is None or ventas_df is None or costes_df is None:
                logger.error(f"No se pudieron leer los datos para la seccion '{seccion}'")
                continue
            
            # Filtrar ventas por semana
            if 'Semana' not in ventas_df.columns:
                if 'Fecha' in ventas_df.columns:
                    ventas_df['Fecha'] = pd.to_datetime(ventas_df['Fecha'], errors='coerce')
                    ventas_df['Semana'] = ventas_df['Fecha'].apply(
                        lambda x: x.isocalendar()[1] if pd.notna(x) else None
                    )
                else:
                    logger.warning(f"No hay columna 'Fecha' ni 'Semana' en ventas de '{seccion}'")
                    continue
            
            datos_semana = ventas_df[ventas_df['Semana'] == semana]
            
            if len(datos_semana) == 0:
                logger.warning(f"No hay datos de ventas para la semana {semana} en '{seccion}'")
                continue
            
            logger.info(f"Datos de ventas: {len(datos_semana)} registros")
            
            # ========================================
            # FASE 1: CALCULAR PEDIDO TEÓRICO
            # ========================================
            
            parametros_seccion = {
                'objetivos_semanales': config.get('secciones', {}).get(seccion, {}).get('objetivos_semanales', {}),
                'objetivo_crecimiento': config.get('parametros', {}).get('objetivo_crecimiento', 0.05),
                'stock_minimo_porcentaje': config.get('parametros', {}).get('stock_minimo_porcentaje', 0.30),
                'festivos': config.get('festivos', {})
            }
            
            pedidos = forecast_engine.calcular_pedido_semana(
                semana, datos_semana, abc_df, costes_df, seccion
            )
            
            if len(pedidos) == 0:
                logger.warning(f"No se generaron pedidos para '{seccion}'")
                continue
            
            # Aplicar stock mínimo dinámico
            pedidos, nuevo_stock, ajustes = forecast_engine.aplicar_stock_minimo(
                pedidos, semana, stock_acumulado
            )
            
            # Actualizar stock acumulado
            stock_acumulado.update(nuevo_stock)
            
            # ========================================
            # FASE 2: APLICAR CORRECCIÓN (si está habilitada)
            # ========================================
            
            if aplicar_correccion:
                pedidos_corregido, metricas = aplicar_correccion_pedido(
                    pedidos.copy(), semana, config,
                    parametros_abc=config.get('parametros', {})
                )
                
                if metricas.get('correccion_aplicada', False):
                    # Guardar métricas por sección
                    metricas_correccion_total[seccion] = metricas
                    
                    # Generar archivo corregido
                    archivo_corregido = generar_archivo_pedido_corregido(
                        pedidos_corregido, semana, seccion, parametros_seccion, config, order_generator
                    )
                    
                    if archivo_corregido:
                        archivos_generados.append(archivo_corregido)
                        logger.info(f"Archivo corregido: {os.path.basename(archivo_corregido)}")
                    
                    # Usar pedido corregido para métricas
                    pedidos_final = pedidos_corregido
                    pedidos_corregidos[seccion] = pedidos_corregido
                else:
                    pedidos_final = pedidos
                    logger.info("Usando pedido teórico (sin corrección)")
            else:
                pedidos_final = pedidos
            
            # ========================================
            # GENERAR ARCHIVO DE SALIDA
            # ========================================
            
            # Generar archivo Excel
            archivo = order_generator.generar_archivo_pedido(pedidos_final, semana, seccion, parametros_seccion)
            
            if archivo:
                archivos_generados.append(archivo)
                
                # Calcular métricas
                if 'Pedido_Final' in pedidos_final.columns:
                    pedidos_validos = pedidos_final[pedidos_final['Pedido_Final'] > 0]
                    articulos = len(pedidos_validos)
                    importe = pedidos_validos['Ventas_Objetivo'].sum()
                else:
                    pedidos_validos = pedidos_final[pedidos_final['Unidades_Pedido'] > 0]
                    articulos = len(pedidos_validos)
                    importe = pedidos_validos['Ventas_Objetivo'].sum()
                
                articulos_totales += articulos
                importe_total += importe
                
                logger.info(f"Archivo generado: {archivo}")
                logger.info(f"  Articulos: {articulos}")
                logger.info(f"  Importe: {importe:.2f}€")
            else:
                logger.warning(f"No se generó archivo para '{seccion}'")
            
            pedidos_totales[seccion] = pedidos_final
            datos_semanales[seccion] = datos_semana
            
        except Exception as e:
            logger.error(f"Error procesando seccion '{seccion}': {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            continue
    
    # Actualizar stock acumulado en el estado
    if stock_acumulado:
        state_manager.actualizar_stock_acumulado(stock_acumulado)
    
    # Generar archivo de resumen si hay datos
    if pedidos_totales:
        resumen_data = []
        for seccion, pedidos in pedidos_totales.items():
            if len(pedidos) > 0:
                resumen_seccion = forecast_engine.generar_resumen_pedido(
                    pedidos, semana, datos_semanales.get(seccion, pd.DataFrame())
                )
                if resumen_seccion:
                    resumen_data.append(resumen_seccion)
        
        if resumen_data:
            resumen_df = pd.DataFrame(resumen_data)
            archivo_resumen = order_generator.generar_resumen_excel(resumen_df, 'VIVERO')
            if archivo_resumen:
                archivos_generados.append(archivo_resumen)
    
    # Registrar ejecución en el estado
    archivo_principal = archivos_generados[0] if archivos_generados else None
    
    notas_correccion = ""
    if metricas_correccion_total:
        articulos_corregidos_total = sum(
            m.get('articulos_corregidos', 0) for m in metricas_correccion_total.values()
        )
        notas_correccion = f" - {articulos_corregidos_total} artículos corregidos en FASE 2"
    
    state_manager.registrar_ejecucion(
        semana=semana,
        archivo_generado=archivo_principal or "Sin archivo",
        articulos=articulos_totales,
        importe=importe_total,
        exitosa=len(archivos_generados) > 0,
        notas=f"Procesadas {len(secciones)} secciones{notas_correccion}"
    )
    
    # ========================================
    # ENVÍO DE EMAILS (si está habilitado)
    # ========================================
    
    resultado_email = {'exito': False, 'razon': 'no_enviado'}
    
    if enviar_email and archivos_generados:
        logger.info("\n" + "=" * 60)
        logger.info("PREPARANDO ENVÍO DE EMAILS")
        logger.info("=" * 60)
        
        # Agrupar archivos por sección
        archivos_por_seccion = agrupar_archivos_por_seccion(archivos_generados, config)
        
        # Enviar emails
        resultado_email = enviar_emails_pedidos(semana, config, archivos_por_seccion)
    
    # ========================================
    # RESUMEN FINAL
    # ========================================
    
    logger.info("\n" + "=" * 70)
    logger.info("RESUMEN DE EJECUCION")
    logger.info("=" * 70)
    logger.info(f"Semana procesada: {semana}")
    logger.info(f"Archivos generados: {len(archivos_generados)}")
    logger.info(f"Total articulos: {articulos_totales}")
    logger.info(f"Total importe: {importe_total:.2f}€")
    
    if metricas_correccion_total:
        logger.info("\nMÉTRICAS DE CORRECCIÓN (FASE 2):")
        logger.info("-" * 40)
        articulos_corregidos_total = 0
        diferencia_total = 0
        for seccion, metricas in metricas_correccion_total.items():
            articulos_corregidos_total += metricas.get('articulos_corregidos', 0)
            diferencia_total += metricas.get('diferencia_unidades', 0)
            logger.info(f"  {seccion}: {metricas.get('articulos_corregidos', 0)} artículos corregidos")
        logger.info(f"  TOTAL: {articulos_corregidos_total} artículos corregidos")
        logger.info(f"  Diferencia neta: {int(diferencia_total):+d} unidades")
    
    # Resumen de emails
    if resultado_email.get('exito'):
        logger.info(f"\n✓ EMAILS ENVIADOS: {resultado_email.get('emails_enviados', 0)}")
    elif resultado_email.get('razon') == 'deshabilitado':
        logger.info("\nEmails deshabilitados")
    elif resultado_email.get('razon') == 'sin_password':
        logger.warning("\n✗ No se enviaron emails (falta configurar EMAIL_PASSWORD)")
    elif not resultado_email.get('exito') and resultado_email.get('emails_fallidos', 0) > 0:
        logger.warning(f"\n✗ EMAILS FALLIDOS: {resultado_email.get('emails_fallidos', 0)}")
    
    logger.info("=" * 70)
    
    return len(archivos_generados) > 0, archivo_principal, articulos_totales, importe_total, metricas_correccion_total, resultado_email


# ============================================
# FUNCIÓN PRINCIPAL
# ============================================

def main():
    """
    Función principal del sistema.
    
    Maneja los argumentos de línea de comandos, coordina la carga de
    configuración y delega la ejecución al proceso correspondiente.
    """
    # Configurar parser de argumentos
    parser = argparse.ArgumentParser(
        description='Sistema de Generación de Pedidos de Compra - Vivero Aranjuez V2 (FASE 1 + FASE 2 + Email)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python main.py                      # Ejecución normal (domingo 15:00)
  python main.py --semana 15          # Forzar semana específica
  python main.py --continuo           # Modo continuo (esperando horario)
  python main.py --status             # Mostrar estado del sistema
  python main.py --reset              # Resetear estado del sistema
  python main.py --semana 15 --sin-correccion    # Solo FASE 1
  python main.py --semana 15 --con-correccion     # FASE 1 + FASE 2 (forzado)
  python main.py --semana 15 --sin-email          # Sin enviar emails
  python main.py --verificar-email                # Verificar configuración de email
        """
    )
    
    parser.add_argument(
        '--semana', '-s',
        type=int,
        help='Número de semana a procesar (para pruebas)'
    )
    parser.add_argument(
        '--continuo', '-c',
        action='store_true',
        help='Ejecutar en modo continuo (esperando el horario programado)'
    )
    parser.add_argument(
        '--status',
        action='store_true',
        help='Mostrar estado del sistema y salir'
    )
    parser.add_argument(
        '--reset',
        action='store_true',
        help='Resetear el estado del sistema'
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Activar logging detallado (DEBUG)'
    )
    parser.add_argument(
        '--log',
        type=str,
        default='logs/sistema.log',
        help='Archivo de log (default: logs/sistema.log)'
    )
    parser.add_argument(
        '--sin-correccion',
        action='store_true',
        help='Ejecutar solo FASE 1 (sin corrección)'
    )
    parser.add_argument(
        '--con-correccion',
        action='store_true',
        help='Forzar ejecución con corrección FASE 2'
    )
    parser.add_argument(
        '--sin-email',
        action='store_true',
        help='No enviar emails después de generar los pedidos'
    )
    parser.add_argument(
        '--verificar-email',
        action='store_true',
        help='Verificar la configuración de email y salir'
    )
    
    args = parser.parse_args()
    
    # Determinar nivel de logging
    nivel_log = logging.DEBUG if args.verbose else logging.INFO
    
    # Configurar logging
    global logger
    logger = configurar_logging(nivel=nivel_log, log_file=args.log)
    
    logger.info("=" * 70)
    logger.info("SISTEMA DE PEDIDOS DE COMPRA - VIVERO ARANJUEZ V2")
    logger.info(f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    # Cargar configuración
    config = cargar_configuracion()
    if config is None:
        logger.error("No se pudo cargar la configuración. Saliendo.")
        sys.exit(1)
    
    # Verificar configuración de email si se solicita
    if args.verificar_email:
        logger.info("\nVERIFICANDO CONFIGURACIÓN DE EMAIL:")
        logger.info("-" * 40)
        
        try:
            email_service = crear_email_service(config)
            verificacion = email_service.verificar_configuracion()
            
            if verificacion['valido']:
                logger.info("✓ Configuración de email válida")
            else:
                logger.warning("✗ Problemas en la configuración:")
                for problema in verificacion['problemas']:
                    logger.warning(f"  - {problema}")
            
            # Verificar variable de entorno
            import os
            if os.environ.get('EMAIL_PASSWORD'):
                logger.info("✓ Variable EMAIL_PASSWORD configurada")
            else:
                logger.warning("✗ Variable EMAIL_PASSWORD NO configurada")
                logger.info("  Para configurarla, ejecuta:")
                logger.info("  export EMAIL_PASSWORD='tu_contraseña'")
            
        except Exception as e:
            logger.error(f"Error al verificar email: {str(e)}")
        
        sys.exit(0)
    
    # Determinar si aplicar corrección
    params_correccion = config.get('parametros_correccion', {})
    correccion_habilitada = params_correccion.get('habilitar_correccion', True)
    aplicar_correccion = correccion_habilitada and not args.sin_correccion
    
    if args.con_correccion:
        aplicar_correccion = True
    
    # Determinar si enviar emails
    email_config = config.get('email', {})
    email_habilitado = email_config.get('habilitar_envio', True)
    enviar_email = email_habilitado and not args.sin_email
    
    logger.info(f"Modo de ejecución: {'FASE 1 + FASE 2' if aplicar_correccion else 'Solo FASE 1'}")
    logger.info(f"Envío de emails: {'Sí' if enviar_email else 'No'}")
    
    # Inicializar StateManager
    state_manager = StateManager(config)
    state_manager.cargar_estado()
    
    # Si se pide resetear el estado
    if args.reset:
        logger.info("Reseteando estado del sistema...")
        state_manager.resetear_estado()
        logger.info("Estado reseteado correctamente.")
        sys.exit(0)
    
    # Si se pide mostrar el estado
    if args.status:
        logger.info("\nESTADO DEL SISTEMA:")
        logger.info(state_manager.obtener_resumen_estado())
        
        scheduler = SchedulerService(config)
        logger.info("\n" + scheduler.simular_proxima_ejecucion())
        
        ultima = state_manager.obtener_ultima_semana_procesada()
        logger.info(f"\nÚltima semana procesada: {ultima if ultima else 'Ninguna'}")
        
        metricas = state_manager.obtener_metricas()
        logger.info(f"Métricas: {metricas}")
        
        sys.exit(0)
    
    # Determinar semana a procesar
    scheduler = SchedulerService(config)
    ultima_procesada = state_manager.obtener_ultima_semana_procesada()
    
    if args.semana:
        # Semana forzada por argumento
        semana = args.semana
        logger.info(f"Semana forzada por argumento: {semana}")
    elif args.continuo:
        # Modo continuo: esperar hasta el horario de ejecución
        logger.info("Modo continuo activado. Esperando horario de ejecución...")
        
        while True:
            es_horario, mensaje = scheduler.verificar_horario_ejecucion()
            if es_horario:
                logger.info("¡Es el horario de ejecución!")
                break
            
            logger.info(mensaje)
            import time
            time.sleep(60)  # Esperar 1 minuto
            
            # Verificar si hay semana pendiente
            semana_a_proc, _ = scheduler.calcular_semana_a_procesar(ultima_procesada)
            if semana_a_proc is None:
                logger.info("No hay semanas pendientes de procesamiento.")
                sys.exit(0)
        
        semana = semana_a_proc
    else:
        # Ejecución normal: verificar horario y calcular semana
        es_horario, mensaje = scheduler.verificar_horario_ejecucion()
        
        if not es_horario:
            logger.warning(f"No es el horario de ejecución: {mensaje}")
            logger.info(scheduler.simular_proxima_ejecucion())
            
            # Verificar si hay semana pendiente
            semana_a_proc, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
            
            if semana_a_proc is None:
                logger.info(msg_semana)
                sys.exit(0)
            
            logger.info(f"pero hay semana pendiente: {msg_semana}")
            logger.info("Use --continuo para esperar hasta el horario de ejecución.")
            sys.exit(0)
        
        # Calcular semana a procesar
        semana, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
        
        if semana is None:
            logger.info(msg_semana)
            sys.exit(0)
        
        logger.info(msg_semana)
    
    # Verificar si la semana ya fue procesada
    if state_manager.verificar_semana_procesada(semana) and not args.semana:
        logger.warning(f"La semana {semana} ya fue procesada anteriormente.")
        logger.info("Use --semana para forzar el reprocesamiento.")
        sys.exit(0)
    
    # Procesar el pedido (FASE 1 + FASE 2 opcional)
    exito, archivo, articulos, importe, metricas_correccion, resultado_email = procesar_pedido_semana(
        semana, config, state_manager, 
        forzar=args.semana is not None,
        aplicar_correccion=aplicar_correccion,
        enviar_email=enviar_email
    )
    
    if exito:
        logger.info(f"\n¡PEDIDO GENERADO EXITOSAMENTE!")
        logger.info(f"Archivo principal: {archivo}")
        logger.info(f"Artículos: {articulos}")
        logger.info(f"Importe: {importe:.2f}€")
        
        if metricas_correccion:
            logger.info(f"\nFASE 2 completada: {len(metricas_correccion)} secciones corregidas")
        
        if resultado_email.get('exito'):
            logger.info(f"\nEmails enviados: {resultado_email.get('emails_enviados', 0)}")
        
        sys.exit(0)
    else:
        logger.error("\nERROR: No se pudo generar el pedido.")
        sys.exit(1)


# ============================================
# PUNTO DE ENTRADA
# ============================================

if __name__ == "__main__":
    main()
