#!/usr/bin/env python3
"""
Módulo StateManager - Persistencia de estado entre ejecuciones

Este módulo gestiona la lectura y escritura del archivo state.json que mantiene
el estado del sistema entre diferentes ejecuciones del script. Almacena información
sobre qué semanas han sido procesadas, el stock acumulado por artículo, y métricas
acumuladas de ejecución.

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-01-31
"""

import json
import os
import logging
import copy
from datetime import datetime
from typing import Optional, Dict, List, Any, Tuple
from pathlib import Path

# Configuración del logger
logger = logging.getLogger(__name__)


class StateManager:
    """
    Gestor del estado persistente del sistema.
    
    Esta clase es responsable de mantener y actualizar el archivo state.json
    que contiene toda la información que debe persistir entre ejecuciones del
    sistema. Esto incluye el tracking de semanas procesadas, el stock acumulado
    por artículo, y métricas de ejecución.
    
    Attributes:
        config (dict): Configuración del sistema
        ruta_archivo (str): Ruta al archivo de estado
        estado (dict): Diccionario con el estado actual cargado
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el StateManager con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.rutas = config.get('rutas', {})
        self.ruta_archivo = self._obtener_ruta_estado()
        self.estado = None
        
        logger.info(f"StateManager inicializado. Archivo de estado: {self.ruta_archivo}")
    
    def _obtener_ruta_estado(self) -> str:
        """
        Obtiene la ruta completa del archivo de estado.
        
        Returns:
            str: Ruta absoluta al archivo state.json
        """
        base = self.rutas.get('directorio_base', '.')
        dir_estado = self.rutas.get('directorio_estado', './data')
        archivo = self.rutas.get('archivo_estado', 'state.json')
        
        # Si es ruta relativa, combinar con base
        if not os.path.isabs(dir_estado):
            dir_estado = os.path.join(base, dir_estado)
        
        return os.path.join(dir_estado, archivo)
    
    def cargar_estado(self) -> Dict[str, Any]:
        """
        Carga el estado desde el archivo JSON.
        
        Si el archivo no existe, crea uno nuevo con la estructura inicial.
        Si existe pero está corrupto, intenta recuperarlo o crea uno nuevo.
        
        Returns:
            Dict[str, Any]: Diccionario con el estado cargado
        """
        logger.info(f"Cargando estado desde: {self.ruta_archivo}")
        
        # Verificar si el archivo existe
        if not os.path.exists(self.ruta_archivo):
            logger.warning(f"Archivo de estado no encontrado. Creando nuevo: {self.ruta_archivo}")
            self.estado = self._crear_estado_inicial()
            self.guardar_estado()
            return self.estado
        
        try:
            with open(self.ruta_archivo, 'r', encoding='utf-8') as f:
                self.estado = json.load(f)
            
            logger.info("Estado cargado correctamente")
            return self.estado
            
        except json.JSONDecodeError as e:
            logger.error(f"Error al parsear estado JSON: {str(e)}")
            logger.warning("Intentando recuperar estado anterior...")
            return self._recuperar_estado()
        
        except Exception as e:
            logger.error(f"Error inesperado al cargar estado: {str(e)}")
            return self._crear_estado_inicial()
    
    def _crear_estado_inicial(self) -> Dict[str, Any]:
        """
        Crea la estructura inicial del estado.
        
        Returns:
            Dict[str, Any]: Estructura inicial del estado
        """
        return {
            "informacion_sistema": {
                "version": "2.0.0",
                "fecha_creacion": datetime.now().isoformat(),
                "ultima_actualizacion": None,
                "ultima_ejecucion_exitosa": None,
                "ultima_semana_procesada": None,
                "año_en_curso": datetime.now().year
            },
            
            "configuracion_actual": {
                "semana_inicio_periodo": None,
                "semana_fin_periodo": None,
                "secciones_procesando": []
            },
            
            "stock_acumulado": {},
            
            "historico_ejecuciones": [],
            
            "pedidos_generados": [],
            
            "metricas": {
                "total_ejecuciones": 0,
                "total_articulos_procesados": 0,
                "total_importe_pedidos": 0.0,
                "ultima_semana_procesada_exitosamente": None
            },
            
            "errores_pendientes": [],
            
            "notas": {
                "es_nota": "Este archivo contiene el estado del sistema entre ejecuciones",
                "instrucciones": "No modificar manualmente a menos que sea necesario para resetear el sistema",
                "para_resetear": "Establecer 'ultima_semana_procesada' a null para reprocesar desde el inicio"
            }
        }
    
    def _recuperar_estado(self) -> Dict[str, Any]:
        """
        Intenta recuperar el estado desde un backup o crea uno nuevo.
        
        Returns:
            Dict[str, Any]: Estado recuperado o nuevo
        """
        # Buscar backup
        backup_path = self.ruta_archivo + '.backup'
        
        if os.path.exists(backup_path):
            try:
                with open(backup_path, 'r', encoding='utf-8') as f:
                    estado_recuperado = json.load(f)
                logger.info(f"Estado recuperado desde backup: {backup_path}")
                return estado_recuperado
            except Exception as e:
                logger.error(f"Error al leer backup: {str(e)}")
        
        # Si no hay backup válido, crear nuevo estado
        logger.warning("No se pudo recuperar el estado. Creando nuevo estado.")
        return self._crear_estado_inicial()
    
    def guardar_estado(self) -> bool:
        """
        Guarda el estado actual en el archivo JSON.
        
        Antes de guardar, crea una copia de seguridad del archivo anterior.
        
        Returns:
            bool: True si se guardó correctamente, False si hubo error
        """
        if self.estado is None:
            logger.error("No hay estado para guardar")
            return False
        
        try:
            # Crear backup del archivo anterior
            if os.path.exists(self.ruta_archivo):
                backup_path = self.ruta_archivo + '.backup'
                try:
                    with open(self.ruta_archivo, 'r', encoding='utf-8') as f:
                        contenido_actual = f.read()
                    with open(backup_path, 'w', encoding='utf-8') as f:
                        f.write(contenido_actual)
                except Exception as e:
                    logger.warning(f"No se pudo crear backup: {str(e)}")
            
            # Guardar nuevo estado
            with open(self.ruta_archivo, 'w', encoding='utf-8') as f:
                json.dump(self.estado, f, indent=4, ensure_ascii=False)
            
            logger.info("Estado guardado correctamente")
            return True
            
        except Exception as e:
            logger.error(f"Error al guardar estado: {str(e)}")
            return False
    
    def obtener_ultima_semana_procesada(self) -> Optional[int]:
        """
        Obtiene el número de la última semana procesada.
        
        Returns:
            Optional[int]: Número de semana o None si no hay ninguna procesada
        """
        if self.estado is None:
            self.cargar_estado()
        
        return self.estado.get('informacion_sistema', {}).get('ultima_semana_procesada')
    
    def establecer_ultima_semana_procesada(self, semana: int) -> bool:
        """
        Actualiza el número de la última semana procesada.
        
        Args:
            semana (int): Número de la semana procesada
        
        Returns:
            bool: True si se actualizó correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        self.estado['informacion_sistema']['ultima_semana_procesada'] = semana
        self.estado['informacion_sistema']['ultima_actualizacion'] = datetime.now().isoformat()
        
        return self.guardar_estado()
    
    def obtener_stock_acumulado(self) -> Dict[str, int]:
        """
        Obtiene el diccionario de stock acumulado por artículo.
        
        Returns:
            Dict[str, int]: Diccionario con clave artículo -> stock
        """
        if self.estado is None:
            self.cargar_estado()
        
        return self.estado.get('stock_acumulado', {})
    
    def actualizar_stock_acumulado(self, articulos_actualizados: Dict[str, int]) -> bool:
        """
        Actualiza el stock acumulado con los valores proporcionados.
        
        Args:
            articulos_actualizados (Dict[str, int]): Diccionario de artículos actualizados
        
        Returns:
            bool: True si se actualizó correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        stock_actual = self.estado.get('stock_acumulado', {})
        stock_actual.update(articulos_actualizados)
        self.estado['stock_acumulado'] = stock_actual
        
        return self.guardar_estado()
    
    def registrar_ejecucion(self, semana: int, archivo_generado: str, 
                            articulos: int, importe: float, exitosa: bool,
                            notas: Optional[str] = None) -> bool:
        """
        Registra una ejecución en el histórico.
        
        Args:
            semana (int): Número de semana procesada
            archivo_generado (str): Nombre del archivo generado
            articulos (int): Número de artículos en el pedido
            importe (float): Importe total del pedido
            exitosa (bool): Si la ejecución fue exitosa
            notas (Optional[str]): Notas adicionales
        
        Returns:
            bool: True si se registró correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        # Crear registro de ejecución
        registro = {
            "semana": semana,
            "fecha_ejecucion": datetime.now().isoformat(),
            "archivo_generado": archivo_generado,
            "num_articulos": articulos,
            "importe": round(importe, 2),
            "exitosa": exitosa,
            "notas": notas or ""
        }
        
        # Agregar al histórico
        historico = self.estado.get('historico_ejecuciones', [])
        historico.append(registro)
        self.estado['historico_ejecuciones'] = historico
        
        # Actualizar métricas
        metricas = self.estado.get('metricas', {})
        metricas['total_ejecuciones'] = metricas.get('total_ejecuciones', 0) + 1
        metricas['total_articulos_procesados'] = metricas.get('total_articulos_procesados', 0) + articulos
        metricas['total_importe_pedidos'] = metricas.get('total_importe_pedidos', 0.0) + importe
        
        if exitosa:
            metricas['ultima_semana_procesada_exitosamente'] = semana
            self.estado['informacion_sistema']['ultima_ejecucion_exitosa'] = datetime.now().isoformat()
            self.estado['informacion_sistema']['ultima_semana_procesada'] = semana
        
        # Agregar a pedidos generados
        pedido = {
            "semana": semana,
            "archivo": archivo_generado,
            "fecha": datetime.now().strftime('%Y-%m-%d'),
            "importe": round(importe, 2)
        }
        pedidos = self.estado.get('pedidos_generados', [])
        pedidos.append(pedido)
        self.estado['pedidos_generados'] = pedidos
        
        return self.guardar_estado()
    
    def obtener_pedidos_por_semana(self, semana: int) -> List[Dict[str, Any]]:
        """
        Obtiene los pedidos generados para una semana específica.
        
        Args:
            semana (int): Número de semana a buscar
        
        Returns:
            List[Dict[str, Any]]: Lista de pedidos de esa semana
        """
        if self.estado is None:
            self.cargar_estado()
        
        pedidos = self.estado.get('pedidos_generados', [])
        return [p for p in pedidos if p.get('semana') == semana]
    
    def verificar_semana_procesada(self, semana: int) -> bool:
        """
        Verifica si una semana ya ha sido procesada.
        
        Args:
            semana (int): Número de semana a verificar
        
        Returns:
            bool: True si la semana ya fue procesada
        """
        ultima = self.obtener_ultima_semana_procesada()
        
        if ultima is None:
            return False
        
        return ultima >= semana
    
    def obtener_metricas(self) -> Dict[str, Any]:
        """
        Obtiene las métricas acumuladas del sistema.
        
        Returns:
            Dict[str, Any]: Diccionario con las métricas
        """
        if self.estado is None:
            self.cargar_estado()
        
        return self.estado.get('metricas', {})
    
    def agregar_error(self, error: Dict[str, Any]) -> bool:
        """
        Agrega un error a la lista de errores pendientes.
        
        Args:
            error (Dict[str, Any]): Información del error
        
        Returns:
            bool: True si se agregó correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        errores = self.estado.get('errores_pendientes', [])
        
        error_registro = {
            "timestamp": datetime.now().isoformat(),
            "tipo": error.get('tipo', 'desconocido'),
            "mensaje": error.get('mensaje', ''),
            "detalles": error.get('detalles', ''),
            "procesado": False
        }
        
        errores.append(error_registro)
        self.estado['errores_pendientes'] = errores
        
        return self.guardar_estado()
    
    def limpiar_errores_procesados(self) -> bool:
        """
        Elimina los errores ya procesados del registro.
        
        Returns:
            bool: True si se limpió correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        errores = self.estado.get('errores_pendientes', [])
        errores_filtrados = [e for e in errores if not e.get('procesado', False)]
        self.estado['errores_pendientes'] = errores_filtrados
        
        return self.guardar_estado()
    
    def resetear_estado(self, mantener_config: bool = True) -> bool:
        """
        Resetea el estado a los valores iniciales.
        
        Args:
            mantener_config (bool): Si True, mantiene la configuración actual
        
        Returns:
            bool: True si se reseteó correctamente
        """
        if self.estado is None:
            self.cargar_estado()
        
        configuracion_actual = self.estado.get('configuracion_actual', {}) if mantener_config else {}
        
        self.estado = self._crear_estado_inicial()
        
        if mantener_config:
            self.estado['configuracion_actual'] = configuracion_actual
        
        return self.guardar_estado()
    
    def obtener_resumen_estado(self) -> str:
        """
        Genera un resumen legible del estado actual del sistema.
        
        Returns:
            str: Resumen formateado del estado
        """
        if self.estado is None:
            self.cargar_estado()
        
        info = self.estado.get('informacion_sistema', {})
        metricas = self.estado.get('metricas', {})
        stock = self.estado.get('stock_acumulado', {})
        pedidos = self.estado.get('pedidos_generados', [])
        
        resumen = []
        resumen.append("=" * 60)
        resumen.append("RESUMEN DEL ESTADO DEL SISTEMA")
        resumen.append("=" * 60)
        resumen.append(f"Versión: {info.get('version', 'N/A')}")
        resumen.append(f"Última semana procesada: {info.get('ultima_semana_procesada', 'Ninguna')}")
        resumen.append(f"Última ejecución exitosa: {info.get('ultima_ejecucion_exitosa', 'N/A')}")
        resumen.append("")
        resumen.append("MÉTRICAS:")
        resumen.append(f"  Total ejecuciones: {metricas.get('total_ejecuciones', 0)}")
        resumen.append(f"  Total artículos procesados: {metricas.get('total_articulos_procesados', 0)}")
        resumen.append(f"  Total importe pedidos: {metricas.get('total_importe_pedidos', 0.0):.2f}€")
        resumen.append("")
        resumen.append(f"STOCK ACUMULADO: {len(stock)} artículos")
        resumen.append(f"PEDIDOS GENERADOS: {len(pedidos)}")
        resumen.append("=" * 60)
        
        return "\n".join(resumen)


# Funciones de utilidad para uso directo
def crear_state_manager(config: dict) -> StateManager:
    """
    Crea una instancia del StateManager.
    
    Args:
        config (dict): Configuración del sistema
    
    Returns:
        StateManager: Instancia inicializada del gestor de estado
    """
    return StateManager(config)


if __name__ == "__main__":
    # Ejemplo de uso
    print("StateManager - Módulo de persistencia de estado")
    print("=" * 50)
    
    # Configurar logging básico
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'rutas': {
            'directorio_base': '.',
            'directorio_estado': './data',
            'archivo_estado': 'state.json'
        }
    }
    
    # Crear StateManager
    sm = crear_state_manager(config_ejemplo)
    
    # Cargar estado
    sm.cargar_estado()
    
    # Mostrar resumen
    print(sm.obtener_resumen_estado())
    
    # Ejemplo de registro de ejecución
    sm.registrar_ejecucion(
        semana=10,
        archivo_generado="Pedido_Semana_10_2026-03-08.xlsx",
        articulos=150,
        importe=12500.50,
        exitosa=True,
        notas="Primera ejecución del sistema V2"
    )
    
    # Verificar semana procesada
    print(f"\nSemana 10 procesada: {sm.verificar_semana_procesada(10)}")
    print(f"Semana 11 procesada: {sm.verificar_semana_procesada(11)}")
