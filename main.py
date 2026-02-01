#!/usr/bin/env python3
"""
Sistema de Pedidos de Compra - Vivero Aranjuez V2

Sistema modular para la generación automática de pedidos de compra
con generación semanal programada y persistencia de estado.

Este es el módulo principal que coordina todos los componentes del sistema:
- DataLoader: Carga de datos de entrada
- StateManager: Persistencia de estado entre ejecuciones
- ForecastEngine: Cálculo de pedidos
- OrderGenerator: Generación de archivos de salida
- SchedulerService: Control de ejecución programada

MODO DE USO:
- Ejecución normal (programada): python main.py
- Ejecución forzada para una semana específica: python main.py --semana 15
- Ejecución en modo continuo (esperando horario): python main.py --continuo
- Mostrar información del estado: python main.py --status

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-01-31
"""

import sys
import os
import json
import logging
import argparse
from datetime import datetime
from typing import Optional, Dict, Any, Tuple

# Añadir el directorio actual al path para imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Imports de los módulos del sistema
from src.data_loader import DataLoader
from src.state_manager import StateManager
from src.forecast_engine import ForecastEngine
from src.order_generator import OrderGenerator
from src.scheduler_service import SchedulerService, EstadoEjecucion

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
# PROCESO PRINCIPAL DE GENERACIÓN DE PEDIDO
# ============================================

def procesar_pedido_semana(semana: int, config: Dict[str, Any], 
                           state_manager: StateManager,
                           forzar: bool = False) -> Tuple[bool, Optional[str], int, float]:
    """
    Procesa el pedido para una semana específica.
    
    Args:
        semana (int): Número de semana a procesar
        config (Dict): Configuración del sistema
        state_manager (StateManager): Gestor de estado
        forzar (bool): Si True, fuerza el procesamiento aunque ya esté hecho
    
    Returns:
        Tuple[bool, Optional[str], int, float]: (éxito, archivo_generado, articulos, importe)
    """
    logger.info("=" * 70)
    logger.info(f"PROCESANDO PEDIDO PARA SEMANA {semana}")
    logger.info("=" * 70)
    
    # Inicializar componentes
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
    datos_semanales = {}  # Almacenar datos_semana para cada sección
    articulos_totales = 0
    importe_total = 0.0
    
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
            
            # Debug: Mostrar columnas disponibles
            logger.debug(f"[DEBUG] Columnas en ventas_df: {list(ventas_df.columns)}")
            
            # Filtrar ventas por semana
            if 'Semana' not in ventas_df.columns:
                # Calcular semana si no existe
                if 'Fecha' in ventas_df.columns:
                    ventas_df['Fecha'] = pd.to_datetime(ventas_df['Fecha'], errors='coerce')
                    ventas_df['Semana'] = ventas_df['Fecha'].apply(
                        lambda x: x.isocalendar()[1] if pd.notna(x) else None
                    )
                else:
                    logger.warning(f"No hay columna 'Fecha' ni 'Semana' en ventas de '{seccion}'")
                    continue
            
            # DEBUG: Mostrar semanas disponibles en los datos
            if 'Semana' in ventas_df.columns:
                semanas_disponibles = ventas_df['Semana'].unique()
                logger.debug(f"[DEBUG] Semanas disponibles en datos: {sorted(semanas_disponibles)}")
                logger.debug(f"[DEBUG] Semana buscando: {semana}")
            
            datos_semana = ventas_df[ventas_df['Semana'] == semana]
            
            logger.debug(f"[DEBUG] datos_semana filtrado: {len(datos_semana)} registros")
            
            if len(datos_semana) == 0:
                logger.warning(f"No hay datos de ventas para la semana {semana} en '{seccion}'")
                continue
            
            logger.info(f"Datos de ventas: {len(datos_semana)} registros")
            
            # Calcular pedido con el motor de forecast
            parametros_seccion = {
                'objetivos_semanales': config.get('secciones', {}).get(seccion, {}).get('objetivos_semanales', {}),
                'objetivo_crecimiento': config.get('parametros', {}).get('objetivo_crecimiento', 0.05),
                'stock_minimo_porcentaje': config.get('parametros', {}).get('stock_minimo_porcentaje', 0.30),
                'festivos': config.get('festivos', {})
            }
            
            logger.debug(f"[DEBUG] Llamando a calcular_pedido_semana con {len(datos_semana)} registros de datos_semana")
            
            pedidos = forecast_engine.calcular_pedido_semana(
                semana, datos_semana, abc_df, costes_df, seccion
            )
            
            logger.debug(f"[DEBUG] Pedidos devueltos: {len(pedidos)} registros")
            if len(pedidos) > 0:
                logger.debug(f"[DEBUG] Columnas en pedidos: {list(pedidos.columns)}")
                logger.debug(f"[DEBUG] Primeras filas de pedidos:\n{pedidos.head()}")
            
            if len(pedidos) == 0:
                logger.warning(f"No se generaron pedidos para '{seccion}'")
                continue
            
            # Aplicar stock mínimo dinámico
            logger.debug(f"[DEBUG] Aplicando stock_minimo a {len(pedidos)} pedidos")
            
            pedidos, nuevo_stock, ajustes = forecast_engine.aplicar_stock_minimo(
                pedidos, semana, stock_acumulado
            )
            
            logger.debug(f"[DEBUG] Tras aplicar_stock_minimo: {len(pedidos)} pedidos")
            if len(pedidos) > 0 and 'Unidades_Pedido' in pedidos.columns:
                articulos_con_pedido = len(pedidos[pedidos['Unidades_Pedido'] > 0])
                logger.debug(f"[DEBUG] Artículos con Unidades_Pedido > 0: {articulos_con_pedido}")
            
            # Actualizar stock acumulado
            stock_acumulado.update(nuevo_stock)
            
            # DEBUG: Mostrar resumen del dataframe antes de generar archivo
            logger.debug(f"[DEBUG] Resumen de pedidos antes de generar archivo:")
            logger.debug(f"  Total registros: {len(pedidos)}")
            if len(pedidos) > 0 and 'Unidades_Pedido' in pedidos.columns:
                pedidos_con_contenido = pedidos[pedidos['Unidades_Pedido'] > 0]
                logger.debug(f"  Registros con Unidades_Pedido > 0: {len(pedidos_con_contenido)}")
                if len(pedidos_con_contenido) > 0:
                    logger.debug(f"  Suma Unidades_Pedido: {pedidos_con_contenido['Unidades_Pedido'].sum()}")
            
            # Generar archivo Excel
            archivo = order_generator.generar_archivo_pedido(pedidos, semana, seccion, parametros_seccion)
            
            if archivo:
                archivos_generados.append(archivo)
                
                # Calcular métricas
                pedidos_validos = pedidos[pedidos['Unidades_Pedido'] > 0]
                articulos = len(pedidos_validos)
                importe = pedidos_validos['Ventas_Objetivo'].sum()
                
                articulos_totales += articulos
                importe_total += importe
                
                logger.info(f"Archivo generado: {archivo}")
                logger.info(f"  Articulos: {articulos}")
                logger.info(f"  Importe: {importe:.2f}€")
            else:
                logger.warning(f"No se generó archivo para '{seccion}'. Verificar datos de entrada.")
            
            pedidos_totales[seccion] = pedidos
            datos_semanales[seccion] = datos_semana  # Guardar datos_semana para uso posterior
            
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
        # Generar resumen consolidado
        resumen_data = []
        for seccion, pedidos in pedidos_totales.items():
            if len(pedidos) > 0:
                resumen_seccion = forecast_engine.generar_resumen_pedido(pedidos, semana, 
                                                                          datos_semanales.get(seccion, pd.DataFrame()))
                if resumen_seccion:
                    resumen_data.append(resumen_seccion)
        
        if resumen_data:
            resumen_df = pd.DataFrame(resumen_data)
            archivo_resumen = order_generator.generar_resumen_excel(resumen_df, 'VIVERO')
            if archivo_resumen:
                archivos_generados.append(archivo_resumen)
    
    # Registrar ejecución en el estado
    archivo_principal = archivos_generados[0] if archivos_generados else None
    state_manager.registrar_ejecucion(
        semana=semana,
        archivo_generado=archivo_principal or "Sin archivo",
        articulos=articulos_totales,
        importe=importe_total,
        exitosa=len(archivos_generados) > 0,
        notas=f"Procesadas {len(secciones)} secciones"
    )
    
    logger.info("\n" + "=" * 70)
    logger.info("RESUMEN DE EJECUCION")
    logger.info("=" * 70)
    logger.info(f"Semana procesada: {semana}")
    logger.info(f"Archivos generados: {len(archivos_generados)}")
    logger.info(f"Total articulos: {articulos_totales}")
    logger.info(f"Total importe: {importe_total:.2f}€")
    logger.info("=" * 70)
    
    return len(archivos_generados) > 0, archivo_principal, articulos_totales, importe_total


# ============================================
# FUNCIÓN PRINCIPAL
# ============================================

def main():
    """
    Función principal del sistema.
    
    Maneja los argumentos de línea de comandos, coordina la carga de
    configuración y delegla ejecución al proceso correspondiente.
    """
    # Configurar parser de argumentos
    parser = argparse.ArgumentParser(
        description='Sistema de Generacion de Pedidos de Compra - Vivero Aranjuez V2',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python main.py                      # Ejecucion normal (domingo 15:00)
  python main.py --semana 15          # Forzar semana especifica
  python main.py --continuo           # Modo continuo (esperando horario)
  python main.py --status             # Mostrar estado del sistema
  python main.py --reset              # Resetear estado del sistema
        """
    )
    
    parser.add_argument(
        '--semana', '-s',
        type=int,
        help='Numero de semana a procesar (para pruebas)'
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
    
    args = parser.parse_args()
    
    # Determinar nivel de logging
    nivel_log = logging.DEBUG if args.verbose else logging.INFO
    
    # Configurar logging
    global logger
    logger = configurar_logging(nivel=nivel_log, log_file=args.log)
    
    logger.info("=" * 70)
    logger.info("SISTEMA DE PEDIDOS DE COMPRA - VIVERO ARANJUEZ V2")
    logger.info(f"Fecha de ejecucion: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    # Cargar configuración
    config = cargar_configuracion()
    if config is None:
        logger.error("No se pudo cargar la configuracion. Saliendo.")
        sys.exit(1)
    
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
        logger.info(f"\nUltima semana procesada: {ultima if ultima else 'Ninguna'}")
        
        metricas = state_manager.obtener_metricas()
        logger.info(f"Metricas: {metricas}")
        
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
        logger.info("Modo continuo activado. Esperando horario de ejecucion...")
        
        while True:
            es_horario, mensaje = scheduler.verificar_horario_ejecucion()
            if es_horario:
                logger.info("¡Es el horario de ejecucion!")
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
            logger.warning(f"No es el horario de ejecucion: {mensaje}")
            logger.info(scheduler.simular_proxima_ejecucion())
            
            # Verificar si hay semana pendiente
            semana_a_proc, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
            
            if semana_a_proc is None:
                logger.info(msg_semana)
                sys.exit(0)
            
            logger.info(f"pero hay semana pendiente: {msg_semana}")
            logger.info("Use --continuo para esperar hasta el horario de ejecucion.")
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
    
    # Procesar el pedido
    exito, archivo, articulos, importe = procesar_pedido_semana(
        semana, config, state_manager, forzar=args.semana is not None
    )
    
    if exito:
        logger.info(f"\n¡PEDIDO GENERADO EXITOSAMENTE!")
        logger.info(f"Archivo: {archivo}")
        logger.info(f"Articulos: {articulos}")
        logger.info(f"Importe: {importe:.2f}€")
        sys.exit(0)
    else:
        logger.error("\nERROR: No se pudo generar el pedido.")
        sys.exit(1)


# ============================================
# PUNTO DE ENTRADA
# ============================================

if __name__ == "__main__":
    main()
