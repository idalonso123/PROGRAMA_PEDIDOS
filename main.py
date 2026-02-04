#!/usr/bin/env python3
"""
Sistema de Pedidos de Compra - Vivero Aranjuez V2

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-04 (Actualizado con correcciones de bugs)
"""

import sys
import os
import json
import logging
import argparse
from datetime import datetime
from typing import Optional, Dict, Any, Tuple, List
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import DataLoader
from src.state_manager import StateManager
from src.forecast_engine import ForecastEngine
from src.order_generator import OrderGenerator
from src.scheduler_service import SchedulerService, EstadoEjecucion

from src.correction_data_loader import CorrectionDataLoader
from src.correction_engine import CorrectionEngine, crear_correction_engine

from src.email_service import EmailService, crear_email_service

import pandas as pd

def configurar_logging(nivel: int = logging.INFO, log_file: Optional[str] = None) -> logging.Logger:
    formato = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    logger = logging.getLogger()
    logger.setLevel(nivel)
    logger.handlers = []
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)
    logger.addHandler(console_handler)
    
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
        file_handler.setFormatter(formato)
        logger.addHandler(file_handler)
    
    return logger

def cargar_configuracion(ruta: str = 'config/config.json') -> Optional[Dict[str, Any]]:
    try:
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

def verificar_archivos_correccion(config: Dict[str, Any], semana: int) -> Dict[str, bool]:
    dir_entrada = config.get('rutas', {}).get('directorio_entrada', './data/input')
    archivos_correccion = config.get('archivos_correccion', {})
    
    disponibilidad = {'stock': False, 'ventas': False, 'compras': False}
    
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
    logger.info("\n" + "=" * 60)
    logger.info("FASE 2: APLICANDO CORRECCIÓN AL PEDIDO")
    logger.info("=" * 60)
    
    params_correccion = config.get('parametros_correccion', {})
    if not params_correccion.get('habilitar_correccion', True):
        logger.info("Corrección deshabilitada en configuración. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False}
    
    disponibilidad = verificar_archivos_correccion(config, semana)
    
    if not any(disponibilidad.values()):
        logger.warning("No se encontraron archivos de corrección. Usando pedido teórico.")
        return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_archivos'}
    
    logger.info(f"Archivos de corrección disponibles: {disponibilidad}")
    
    try:
        correction_loader = CorrectionDataLoader(config)
        datos_correccion = correction_loader.cargar_datos_correccion(semana)
        
        datos_cargados = sum(1 for v in datos_correccion.values() if v is not None)
        if datos_cargados == 0:
            logger.warning("No se pudieron cargar datos de corrección. Usando pedido teórico.")
            return pedido_teorico.copy(), {'correccion_aplicada': False, 'razon': 'sin_datos'}
        
        pedido_fusionado = correction_loader.merge_con_pedido_teorico(
            pedido_teorico, datos_correccion
        )
        
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
        
        metricas = engine.calcular_metricas_correccion(
            pedido_corregido,
            columna_pedido_original='Unidades_Pedido',
            columna_pedido_corregido='Pedido_Corregido',
            columna_ventas_reales='Unidades_Vendidas',
            columna_ventas_objetivo='Ventas_Objetivo'
        )
        metricas['correccion_aplicada'] = True
        metricas['datos_cargados'] = datos_cargados
        
        alertas = engine.generar_alertas(pedido_corregido)
        if alertas:
            metricas['alertas'] = alertas
            logger.warning("ALERTAS GENERADAS:")
            for alerta in alertas:
                logger.warning(f"  [{alerta['nivel']}] {alerta['mensaje']}")
        
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
    try:
        from datetime import datetime, timedelta
        fecha_base = datetime.now()
        
        dia_semana = fecha_base.weekday()
        dias_hasta_lunes = (7 - dia_semana) % 7
        fecha_lunes = fecha_base + timedelta(days=dias_hasta_lunes + (7 * ((semana - fecha_base.isocalendar()[1]) % 52)))
        
        if fecha_lunes < datetime.now():
            fecha_lunes = datetime.now()
        
        fecha_lunes_str = fecha_lunes.strftime('%Y-%m-%d')
        
        dir_salida = config.get('rutas', {}).get('directorio_salida', './data/output')
        nombre_archivo = f"Pedido_Semana_{semana}_{fecha_lunes_str}_{seccion}_CORREGIDO.xlsx"
        ruta_archivo = os.path.join(dir_salida, nombre_archivo)
        
        df_exportar = pedido_corregido.copy()
        
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
        
        columnas_orden = [
            'Código artículo', 'Nombre artículo', 'Talla', 'Color', 'Categoria',
            'Pedido_Teorico', 'Stock_Minimo', 'Stock_Real', 'Ajuste_Stock',
            'Pedido_Final', 'Correccion_Aplicada', 'Escenario',
            'Ventas_Reales', 'Ventas_Objetivo', 'Compras_Recibidas',
            'PVP', 'Coste', 'Proveedor', 'Unidades_ABC', 'Ventas_Objetivo'
        ]
        
        columnas_finales = [col for col in columnas_orden if col in df_exportar.columns]
        df_exportar = df_exportar[columnas_finales]
        
        df_exportar.to_excel(ruta_archivo, index=False, sheet_name=seccion.capitalize())
        
        logger.info(f"Archivo de pedido corregido generado: {nombre_archivo}")
        return ruta_archivo
        
    except Exception as e:
        logger.error(f"Error al generar archivo corregido: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return None

def enviar_emails_pedidos(
    semana: int,
    config: Dict[str, Any],
    archivos_por_seccion: Dict[str, List[str]]
) -> Dict[str, Any]:
    logger.info("\n" + "=" * 60)
    logger.info("ENVÍO DE EMAILS A RESPONSABLES")
    logger.info("=" * 60)
    
    email_config = config.get('email', {})
    if not email_config.get('habilitar_envio', True):
        logger.info("Envío de emails deshabilitado en configuración.")
        return {'exito': False, 'razon': 'deshabilitado'}
    
    try:
        email_service = crear_email_service(config)
        
        verificacion = email_service.verificar_configuracion()
        if not verificacion['valido']:
            logger.warning("Problemas en la configuración de email:")
            for problema in verificacion['problemas']:
                logger.warning(f"  - {problema}")
            
            if any('EMAIL_PASSWORD' in p for p in verificacion['problemas']):
                logger.error("No se puede enviar emails sin configurar la variable EMAIL_PASSWORD")
                return {'exito': False, 'razon': 'sin_password'}
        
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
    CORRECCIÓN: Esta función ahora maneja correctamente secciones con guiones bajos.
    
    El problema original era que para archivos como:
    - Pedido_Semana_14_2026-02-03_mascotas_vivo.xlsx
    - Pedido_Semana_14_2026-02-03_deco_interior.xlsx
    
    El código dividía por '_' y tomaba solo la última parte, convirtiendo:
    - 'mascotas_vivo' en 'vivo'
    - 'deco_interior' en 'interior'
    
    La solución es reconstruir la sección uniendo todas las partes después de la fecha.
    """
    archivos_por_seccion = {}
    dir_salida = config.get('rutas', {}).get('directorio_salida', './data/output')

    for archivo in archivos_generados:
        if not archivo:
            continue

        nombre_archivo = os.path.basename(archivo)
        if 'RESUMEN' in nombre_archivo.upper():
            continue

        nombre_sin_extension = nombre_archivo.replace('.xlsx', '')
        
        partes = nombre_sin_extension.split('_')
        
        if len(partes) >= 4:
            # La fecha es la parte 3 (formato YYYY-MM-DD con guiones)
            # Comprobamos si la parte 3 contiene '-' para verificar si es fecha
            if '-' in partes[3]:
                # CORRECCIÓN: La sección es todo lo que viene después de la fecha
                # Unimos las partes desde el índice 4 en adelante con '_'
                # Esto preserva secciones como 'mascotas_vivo' y 'deco_interior'
                seccion = '_'.join(partes[4:])
            else:
                # Si por alguna razón la fecha no está en su posición,
                # usamos el comportamiento original (última parte)
                seccion = partes[-1]

            if seccion not in archivos_por_seccion:
                archivos_por_seccion[seccion] = []

            archivos_por_seccion[seccion].append(archivo)

    logger.debug(f"[DEBUG] Archivos agrupados por sección: {archivos_por_seccion}")

    return archivos_por_seccion

def procesar_pedido_semana(
    semana: int, 
    config: Dict[str, Any], 
    state_manager: StateManager,
    forzar: bool = False,
    aplicar_correccion: bool = True,
    enviar_email: bool = True
) -> Tuple[bool, Optional[str], int, float, Dict[str, Any], Dict[str, Any]]:
    logger.info("=" * 70)
    logger.info(f"PROCESANDO PEDIDO PARA SEMANA {semana}")
    logger.info("=" * 70)
    
    if aplicar_correccion:
        logger.info("MODO: FASE 1 (Forecast) + FASE 2 (Corrección)")
    else:
        logger.info("MODO: Solo FASE 1 (Forecast) - Corrección deshabilitada")
    
    data_loader = DataLoader(config)
    forecast_engine = ForecastEngine(config)
    order_generator = OrderGenerator(config)
    scheduler = SchedulerService(config)
    
    fecha_lunes, fecha_domingo, fecha_archivo = scheduler.calcular_fechas_semana_pedido(semana)
    logger.info(f"Período de la semana: {fecha_lunes} al {fecha_domingo}")
    
    stock_acumulado = state_manager.obtener_stock_acumulado()
    logger.info(f"Stock acumulado cargado: {len(stock_acumulado)} artículos")
    
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
            abc_df, ventas_df, costes_df = data_loader.leer_datos_seccion(seccion)
            
            logger.debug(f"[DEBUG] abc_df: {len(abc_df) if abc_df is not None else 0} registros")
            logger.debug(f"[DEBUG] ventas_df: {len(ventas_df) if ventas_df is not None else 0} registros")
            logger.debug(f"[DEBUG] costes_df: {len(costes_df) if costes_df is not None else 0} registros")
            
            if abc_df is None or ventas_df is None or costes_df is None:
                logger.error(f"No se pudieron leer los datos para la seccion '{seccion}'")
                continue
            
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
            
            pedidos, nuevo_stock, ajustes = forecast_engine.aplicar_stock_minimo(
                pedidos, semana, stock_acumulado
            )
            
            stock_acumulado.update(nuevo_stock)
            
            if aplicar_correccion:
                pedidos_corregido, metricas = aplicar_correccion_pedido(
                    pedidos.copy(), semana, config,
                    parametros_abc=config.get('parametros', {})
                )
                
                if metricas.get('correccion_aplicada', False):
                    metricas_correccion_total[seccion] = metricas
                    
                    archivo_corregido = generar_archivo_pedido_corregido(
                        pedidos_corregido, semana, seccion, parametros_seccion, config, order_generator
                    )
                    
                    if archivo_corregido:
                        archivos_generados.append(archivo_corregido)
                        logger.info(f"Archivo corregido: {os.path.basename(archivo_corregido)}")
                    
                    pedidos_final = pedidos_corregido
                    pedidos_corregidos[seccion] = pedidos_corregido
                else:
                    pedidos_final = pedidos
                    logger.info("Usando pedido teórico (sin corrección)")
            else:
                pedidos_final = pedidos
            
            archivo = order_generator.generar_archivo_pedido(pedidos_final, semana, seccion, parametros_seccion)
            
            if archivo:
                archivos_generados.append(archivo)
                
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
    
    if stock_acumulado:
        state_manager.actualizar_stock_acumulado(stock_acumulado)
    
    # CORRECCIÓN: Generar archivo de resumen para CADA SECCIÓN y uno consolidado
    if pedidos_totales:
        resumen_data = []
        for seccion, pedidos in pedidos_totales.items():
            if len(pedidos) > 0:
                # CORRECCIÓN: Pasar la sección correcta a generar_resumen_pedido
                resumen_seccion = forecast_engine.generar_resumen_pedido(
                    pedidos, semana, datos_semanales.get(seccion, pd.DataFrame()), seccion
                )
                if resumen_seccion:
                    resumen_data.append(resumen_seccion)
        
        if resumen_data:
            resumen_df = pd.DataFrame(resumen_data)
            # CORRECCIÓN: Generar un resumen consolidado con TODAS las secciones
            archivo_resumen = order_generator.generar_resumen_excel(resumen_df, 'CONSOLIDADO')
            if archivo_resumen:
                archivos_generados.append(archivo_resumen)
    
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
    
    resultado_email = {'exito': False, 'razon': 'no_enviado'}
    
    if enviar_email and archivos_generados:
        logger.info("\n" + "=" * 60)
        logger.info("PREPARANDO ENVÍO DE EMAILS")
        logger.info("=" * 60)
        
        # CORRECCIÓN: Usar la función grouping corregida
        archivos_por_seccion = agrupar_archivos_por_seccion(archivos_generados, config)
        
        resultado_email = enviar_emails_pedidos(semana, config, archivos_por_seccion)
    
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

def main():
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
    
    parser.add_argument('--semana', '-s', type=int, help='Número de semana a procesar (para pruebas)')
    parser.add_argument('--continuo', '-c', action='store_true', help='Ejecutar en modo continuo')
    parser.add_argument('--status', action='store_true', help='Mostrar estado del sistema y salir')
    parser.add_argument('--reset', action='store_true', help='Resetear el estado del sistema')
    parser.add_argument('--verbose', '-v', action='store_true', help='Activar logging detallado (DEBUG)')
    parser.add_argument('--log', type=str, default='logs/sistema.log', help='Archivo de log (default: logs/sistema.log)')
    parser.add_argument('--sin-correccion', action='store_true', help='Ejecutar solo FASE 1 (sin corrección)')
    parser.add_argument('--con-correccion', action='store_true', help='Forzar ejecución con corrección FASE 2')
    parser.add_argument('--sin-email', action='store_true', help='No enviar emails después de generar los pedidos')
    parser.add_argument('--verificar-email', action='store_true', help='Verificar la configuración de email y salir')
    
    args = parser.parse_args()
    
    nivel_log = logging.DEBUG if args.verbose else logging.INFO
    
    global logger
    logger = configurar_logging(nivel=nivel_log, log_file=args.log)
    
    logger.info("=" * 70)
    logger.info("SISTEMA DE PEDIDOS DE COMPRA - VIVERO ARANJUEZ V2")
    logger.info(f"Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)
    
    config = cargar_configuracion()
    if config is None:
        logger.error("No se pudo cargar la configuración. Saliendo.")
        sys.exit(1)
    
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
    
    params_correccion = config.get('parametros_correccion', {})
    correccion_habilitada = params_correccion.get('habilitar_correccion', True)
    aplicar_correccion = correccion_habilitada and not args.sin_correccion
    
    if args.con_correccion:
        aplicar_correccion = True
    
    email_config = config.get('email', {})
    email_habilitado = email_config.get('habilitar_envio', True)
    enviar_email = email_habilitado and not args.sin_email
    
    logger.info(f"Modo de ejecución: {'FASE 1 + FASE 2' if aplicar_correccion else 'Solo FASE 1'}")
    logger.info(f"Envío de emails: {'Sí' if enviar_email else 'No'}")
    
    state_manager = StateManager(config)
    state_manager.cargar_estado()
    
    if args.reset:
        logger.info("Reseteando estado del sistema...")
        state_manager.resetear_estado()
        logger.info("Estado reseteado correctamente.")
        sys.exit(0)
    
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
    
    scheduler = SchedulerService(config)
    ultima_procesada = state_manager.obtener_ultima_semana_procesada()
    
    if args.semana:
        semana = args.semana
        logger.info(f"Semana forzada por argumento: {semana}")
    elif args.continuo:
        logger.info("Modo continuo activado. Esperando horario de ejecución...")
        
        while True:
            es_horario, mensaje = scheduler.verificar_horario_ejecucion()
            if es_horario:
                logger.info("¡Es el horario de ejecución!")
                break
            
            logger.info(mensaje)
            import time
            time.sleep(60)
            
            semana_a_proc, _ = scheduler.calcular_semana_a_procesar(ultima_procesada)
            if semana_a_proc is None:
                logger.info("No hay semanas pendientes de procesamiento.")
                sys.exit(0)
        
        semana = semana_a_proc
    else:
        es_horario, mensaje = scheduler.verificar_horario_ejecucion()
        
        if not es_horario:
            logger.warning(f"No es el horario de ejecución: {mensaje}")
            logger.info(scheduler.simular_proxima_ejecucion())
            
            semana_a_proc, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
            
            if semana_a_proc is None:
                logger.info(msg_semana)
                sys.exit(0)
            
            logger.info(f"pero hay semana pendiente: {msg_semana}")
            logger.info("Use --continuo para esperar hasta el horario de ejecución.")
            sys.exit(0)
        
        semana, msg_semana = scheduler.calcular_semana_a_procesar(ultima_procesada)
        
        if semana is None:
            logger.info(msg_semana)
            sys.exit(0)
        
        logger.info(msg_semana)
    
    if state_manager.verificar_semana_procesada(semana) and not args.semana:
        logger.warning(f"La semana {semana} ya fue procesada anteriormente.")
        logger.info("Use --semana para forzar el reprocesamiento.")
        sys.exit(0)
    
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

if __name__ == "__main__":
    main()
