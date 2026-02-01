#!/usr/bin/env python3
"""
Módulo SchedulerService - Control de ejecución programada

Este módulo gestiona la lógica de programación y control de ejecución del
sistema. Determina cuándo debe ejecutarse el proceso de generación de pedidos
según la configuración de horario (domingos a las 15:00), identifica qué semana
procesar, y verifica si la semana ya ha sido procesada anteriormente.

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-01-31
"""

import logging
import os
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, Tuple
from enum import Enum

# Configuración del logger
logger = logging.getLogger(__name__)


class EstadoEjecucion(Enum):
    """Enumeración de los posibles estados de ejecución."""
    PENDIENTE = "pendiente"
    EJECUTANDO = "ejecutando"
    COMPLETADA = "completada"
    ERROR = "error"
    YA_PROCESADA = "ya_procesada"


class SchedulerService:
    """
    Servicio de control de ejecución programada.
    
    Esta clase encapsula toda la lógica relacionada con el control temporal
    del sistema: verificación de horarios de ejecución, cálculo de semanas
    a procesar, y gestión del estado de ejecuciones.
    
    Attributes:
        config (dict): Configuración del sistema
        horario (dict): Configuración de horario de ejecución
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el SchedulerService con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.horario = config.get('horario_ejecucion', {})
        self.secciones = config.get('secciones_activas', [])
        
        # Valores por defecto si no están configurados
        self.dia_ejecucion = self.horario.get('dia', 'sunday')
        self.hora_ejecucion = self.horario.get('hora', 15)
        self.minuto_ejecucion = self.horario.get('minuto', 0)
        
        logger.info("SchedulerService inicializado correctamente")
        logger.info(f"Horario configurado: {self.dia_ejecucion} a las {self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d}")
    
    def obtener_dia_semana_ingles(self) -> str:
        """
        Obtiene el nombre del día actual en inglés (formato para datetime).
        
        Returns:
            str: Nombre del día en inglés ('monday', 'tuesday', etc.)
        """
        dias = {
            'monday': 'monday',
            'tuesday': 'tuesday',
            'wednesday': 'wednesday',
            'thursday': 'thursday',
            'friday': 'friday',
            'saturday': 'saturday',
            'sunday': 'sunday',
            'lunes': 'monday',
            'martes': 'tuesday',
            'miercoles': 'wednesday',
            'jueves': 'thursday',
            'viernes': 'friday',
            'sabado': 'saturday',
            'domingo': 'sunday'
        }
        
        dia_config = self.dia_ejecucion.lower()
        return dias.get(dia_config, 'sunday')
    
    def obtener_numero_semana_actual(self) -> int:
        """
        Obtiene el número de semana ISO actual.
        
        Returns:
            int: Número de semana ISO (1-53)
        """
        return datetime.now().isocalendar()[1]
    
    def obtener_numero_semana_siguiente(self) -> int:
        """
        Obtiene el número de la siguiente semana (semana actual + 1).
        
        Returns:
            int: Número de la siguiente semana ISO
        """
        semana_actual = self.obtener_numero_semana_actual()
        semana_siguiente = semana_actual + 1
        
        # Si pasamos de semana 52/53 a semana 1 del siguiente año
        if semana_siguiente > 53:
            semana_siguiente = 1
        
        return semana_siguiente
    
    def obtener_fecha_actual(self) -> datetime:
        """
        Obtiene la fecha y hora actual del sistema.
        
        Returns:
            datetime: Fecha y hora actual
        """
        return datetime.now()
    
    def verificar_horario_ejecucion(self) -> Tuple[bool, str]:
        """
        Verifica si la fecha y hora actual coincide con el horario de ejecución.
        
        Returns:
            Tuple[bool, str]: (es_horario_correcto, mensaje_explicativo)
        """
        ahora = self.obtener_fecha_actual()
        dia_actual = ahora.strftime('%A').lower()
        hora_actual = ahora.hour
        minuto_actual = ahora.minuto
        
        # Verificar día de la semana
        dia_ejecucion = self.obtener_dia_semana_ingles()
        
        if dia_actual != dia_ejecucion:
            dias_es = {
                'monday': 'lunes',
                'tuesday': 'martes',
                'wednesday': 'miércoles',
                'thursday': 'jueves',
                'friday': 'viernes',
                'saturday': 'sábado',
                'sunday': 'domingo'
            }
            return False, f"Hoy es {dias_es.get(dia_actual, dia_actual)}, no es {dias_es.get(dia_ejecucion, dia_ejecucion)}"
        
        # Verificar hora de ejecución
        if hora_actual < self.hora_ejecucion:
            return False, f"La hora actual ({hora_actual:02d}:{minuto_actual:02d}) es anterior a la hora de ejecución ({self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d})"
        
        # Verificar si ya pasó demasiado tiempo de la ventana de ejecución
        if hora_actual > self.hora_ejecucion + 1:
            return False, f"La hora actual ({hora_actual:02d}:{minuto_actual:02d}) ya pasó la ventana de ejecución"
        
        return True, "Es el horario correcto para ejecutar"
    
    def verificar_ejecucion_semana(self, semana: int, ultima_semana_procesada: Optional[int]) -> Tuple[EstadoEjecucion, str]:
        """
        Verifica el estado de ejecución para una semana específica.
        
        Args:
            semana (int): Número de semana a verificar
            ultima_semana_procesada (Optional[int]): Última semana procesada según el estado
        
        Returns:
            Tuple[EstadoEjecucion, str]: (estado, mensaje_explicativo)
        """
        # Si no hay semana procesada anteriormente, ejecutar
        if ultima_semana_procesada is None:
            return EstadoEjecucion.PENDIENTE, f"La semana {semana} no ha sido procesada anteriormente"
        
        # Si la semana ya fue procesada, no volver a procesar
        if ultima_semana_procesada >= semana:
            return EstadoEjecucion.YA_PROCESADA, f"La semana {semana} ya fue procesada (última procesada: {ultima_semana_procesada})"
        
        # Si hay semanas pendientes entre la última procesada y la actual, ejecutar
        return EstadoEjecucion.PENDIENTE, f"La semana {semana} está pendiente de procesamiento"
    
    def calcular_semana_a_procesar(self, ultima_semana_procesada: Optional[int],
                                     forzar_semana: Optional[int] = None) -> Tuple[Optional[int], str]:
        """
        Calcula qué semana debe procesarse según la configuración y el estado.
        
        Args:
            ultima_semana_procesada (Optional[int]): Última semana procesada
            forzar_semana (Optional[int]): Semana específica a forzar (para pruebas)
        
        Returns:
            Tuple[Optional[int], str]: (semana_a_procesar, mensaje)
        """
        # Si se fuerza una semana específica (para pruebas)
        if forzar_semana is not None:
            return forzar_semana, f"Semana forzada: {forzar_semana}"
        
        # Calcular la semana siguiente a la última procesada
        if ultima_semana_procesada is None:
            semana_siguiente = self.obtener_numero_semana_actual()
        else:
            semana_siguiente = ultima_semana_procesada + 1
        
        # Verificar límites de semanas
        config = self.config.get('parametros', {})
        semana_inicio = config.get('semana_inicio', 1)
        semana_fin = config.get('semana_fin', 52)
        
        if semana_siguiente < semana_inicio:
            return None, f"La semana {semana_siguiente} está fuera del período ({semana_inicio}-{semana_fin})"
        
        if semana_siguiente > semana_fin:
            return None, f"La semana {semana_siguiente} está fuera del período ({semana_inicio}-{semana_fin})"
        
        return semana_siguiente, f"Semana a procesar: {semana_siguiente}"
    
    def calcular_fechas_semana_pedido(self, semana: int, año: Optional[int] = None) -> Tuple[str, str, str]:
        """
        Calcula las fechas relevantes para el pedido de una semana.
        
        Args:
            semana (int): Número de semana ISO
            año (Optional[int]): Año (usa el actual si no se especifica)
        
        Returns:
            Tuple[str, str, str]: (fecha_lunes, fecha_domingo, fecha_formateada)
        """
        if año is None:
            año = datetime.now().year
        
        # Calcular fechas de la semana
        from datetime import date, timedelta
        
        # Obtener el jueves de la semana (día central de la semana ISO)
        fecha_base = date(año, 1, 4)  # 4 de enero siempre está en la semana 1 del año ISO
        delta = timedelta(weeks=semana - 1)
        fecha = fecha_base + delta
        
        # Obtener el lunes de esa semana
        dias_lunes = fecha.weekday()
        lunes = fecha - timedelta(days=dias_lunes)
        
        # El fin de semana es el domingo
        domingo = lunes + timedelta(days=6)
        
        # Fecha formateada para文件名
        fecha_formateada = datetime.now().strftime('%Y-%m-%d')
        
        return lunes.strftime('%Y-%m-%d'), domingo.strftime('%Y-%m-%d'), fecha_formateada
    
    def obtener_dias_hasta_ejecucion(self) -> int:
        """
        Calcula los días que faltan hasta la próxima ejecución programada.
        
        Returns:
            int: Días hasta la próxima ejecución (0 si es hoy)
        """
        ahora = datetime.now()
        dia_ejecucion = self.obtener_dia_semana_ingles()
        
        # Días de la semana (0 = lunes, 6 = domingo)
        dias_map = {
            'monday': 0,
            'tuesday': 1,
            'wednesday': 2,
            'thursday': 3,
            'friday': 4,
            'saturday': 5,
            'sunday': 6
        }
        
        dia_actual = ahora.weekday()
        dia_target = dias_map.get(dia_ejecucion, 6)
        
        # Calcular días hasta el próximo día de ejecución
        dias_hasta = (dia_target - dia_actual) % 7
        
        return dias_hasta
    
    def es_modo_prueba(self) -> bool:
        """
        Verifica si el sistema está en modo prueba.
        
        Returns:
            bool: True si está en modo prueba
        """
        return os.environ.get('MODO_PRUEBA', 'false').lower() == 'true'
    
    def obtener_resumen_estado(self) -> Dict[str, Any]:
        """
        Genera un resumen del estado del scheduler.
        
        Returns:
            Dict: Resumen con información del estado actual
        """
        ahora = datetime.now()
        
        resumen = {
            'fecha_actual': ahora.strftime('%Y-%m-%d %H:%M:%S'),
            'semana_actual': self.obtener_numero_semana_actual(),
            'proxima_ejecucion': {
                'dia': self.dia_ejecucion,
                'hora': f"{self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d}",
                'dias_falta': self.obtener_dias_hasta_ejecucion()
            },
            'configuracion': {
                'secciones_configuradas': len(self.secciones),
                'secciones': self.secciones
            },
            'modo_prueba': self.es_modo_prueba()
        }
        
        return resumen
    
    def simular_proxima_ejecucion(self) -> str:
        """
        Genera un mensaje explicando cuándo será la próxima ejecución.
        
        Returns:
            str: Mensaje formateado con la próxima fecha de ejecución
        """
        dias_map = {
            0: 'lunes',
            1: 'martes',
            2: 'miércoles',
            3: 'jueves',
            4: 'viernes',
            5: 'sábado',
            6: 'domingo'
        }
        
        dias_hasta = self.obtener_dias_hasta_ejecucion()
        ahora = datetime.now()
        
        if dias_hasta == 0:
            if self.es_modo_prueba():
                return f"MODO PRUEBA: La ejecución puede realizarse ahora (domingo a las {self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d})"
            else:
                # Verificar si ya pasó la hora de ejecución
                if ahora.hour >= self.hora_ejecucion:
                    return f"La ventana de ejecución de hoy ({self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d}) ya pasó"
                else:
                    return f"La ejecución está programada para hoy a las {self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d}"
        else:
            proxima_fecha = ahora + timedelta(days=dias_hasta)
            nombre_dia = dias_map.get(proxima_fecha.weekday(), 'domingo')
            return f"La próxima ejecución será el {nombre_dia} {proxima_fecha.strftime('%d/%m/%Y')} a las {self.hora_ejecucion:02d}:{self.minuto_ejecucion:02d}"


# Funciones de utilidad para uso directo
def crear_scheduler_service(config: dict) -> SchedulerService:
    """
    Crea una instancia del SchedulerService.
    
    Args:
        config (dict): Configuración del sistema
    
    Returns:
        SchedulerService: Instancia inicializada del servicio de scheduler
    """
    return SchedulerService(config)


if __name__ == "__main__":
    # Ejemplo de uso
    print("SchedulerService - Módulo de control de ejecución")
    print("=" * 50)
    
    # Configurar logging básico
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'horario_ejecucion': {
            'dia': 'sunday',
            'hora': 15,
            'minuto': 0
        },
        'secciones_activas': ['vivero', 'interior', 'maf'],
        'parametros': {
            'semana_inicio': 10,
            'semana_fin': 22
        }
    }
    
    # Crear SchedulerService
    scheduler = crear_scheduler_service(config_ejemplo)
    
    # Mostrar estado actual
    print("\nEstado actual del scheduler:")
    print(scheduler.simular_proxima_ejecucion())
    
    # Mostrar resumen completo
    print("\nResumen del scheduler:")
    resumen = scheduler.obtener_resumen_estado()
    for clave, valor in resumen.items():
        print(f"  {clave}: {valor}")
    
    # Verificar horario
    print("\nVerificación de horario:")
    es_horario, mensaje = scheduler.verificar_horario_ejecucion()
    print(f"  ¿Es horario correcto?: {es_horario}")
    print(f"  Mensaje: {mensaje}")
    
    # Ejemplo de verificación de semana
    print("\nVerificación de semana:")
    estado, mensaje = scheduler.verificar_ejecucion_semana(15, 14)
    print(f"  Estado: {estado.value}")
    print(f"  Mensaje: {mensaje}")
