#!/usr/bin/env python3
"""
Módulo FileFinder - Búsqueda flexible de archivos con sufijos dinámicos

Este módulo proporciona funciones para buscar archivos generados por el ERP
que incluyen timestamps dinámicos en su nombre. El formato del ERP es:

    BASE__YYYYMMDD_HHMMSS.ext

Donde:
- BASE: Nombre base del archivo (ej: SPA_Ventas, SPA_Stock_actual)
- __: Separador doble guión bajo
- YYYYMMDD: Fecha de exportación (8 dígitos)
- HHMMSS: Hora de exportación (6 dígitos)
- .ext: Extensión del archivo

La función principal `find_latest_file()` busca automáticamente el archivo
más reciente cuando hay múltiples exportaciones del mismo día.

Autor: Sistema de Pedidos V2
Fecha: 2026-02-06
"""

import os
import glob
import re
from pathlib import Path
from typing import Optional, List
import logging

# Configuración del logger
logger = logging.getLogger(__name__)


def find_latest_file(
    directory: str, 
    base_name: str, 
    extension: str = ".xlsx",
    legacy_fallback: bool = True,
    logger_instance: Optional[logging.Logger] = None
) -> Optional[str]:
    """
    Busca archivos que coincidan con el patrón BASE__YYYYMMDD_HHMMSS.EXT
    y devuelve el más reciente (mayor timestamp).
    
    Esta función está diseñada para manejar los archivos exportados por el ERP,
    que incluyen un timestamp en el nombre. Cuando hay múltiples exportaciones
    del mismo archivo (por ejemplo, exports de mañana y tarde), esta función
    selecciona automáticamente el más reciente.
    
    Args:
        directory: Directorio donde buscar los archivos
        base_name: Nombre base del archivo (sin extensión, sin timestamp)
        extension: Extensión del archivo (default: ".xlsx")
        legacy_fallback: Si True, busca archivo legacy (BASE.EXT) si no hay timestamped
        logger_instance: Logger personalizado (usa el global si no se proporciona)
    
    Returns:
        Ruta completa del archivo más reciente encontrado, o None si no existe ninguno
    
    Examples:
        >>> find_latest_file("./data/input", "SPA_Ventas")
        '/workspace/data/input/SPA_Ventas__20260205_210037.xlsx'
        
        >>> find_latest_file("./data/input", "Stock", legacy_fallback=True)
        '/workspace/data/input/Stock.xlsx'  # Si no hay archivo con timestamp
    
    Raises:
        FileNotFoundError: Si no se encuentra ningún archivo (ni timestamped ni legacy)
    """
    # Usar logger proporcionado o el global
    log = logger_instance if logger_instance else logger
    
    # Normalizar separadores de ruta
    directory = os.path.normpath(directory)
    base_name = base_name.strip()
    
    # Verificar que el directorio existe
    if not os.path.isdir(directory):
        log.warning(f"Directorio no existe: {directory}")
        return None
    
    # PATRÓN PRINCIPAL: BASE__????????_??????.EXT
    # Los signos ? representan exactamente 1 dígito
    # El patrón busca: nombre_base__ seguido de 8 dígitos (fecha), _, 6 dígitos (hora), extensión
    timestamp_pattern = f"{base_name}__????????_??????{extension}"
    full_pattern = os.path.join(directory, timestamp_pattern)
    
    # Buscar todos los archivos que coincidan con el patrón de timestamp
    matching_files = glob.glob(full_pattern)
    
    if matching_files:
        # Ordenar por nombre de archivo (el timestamp está en formato ISO-like)
        # Los archivos más recientes tienen timestamps mayores alfabéticamente
        matching_files.sort(reverse=True)
        latest_file = matching_files[0]
        log.info(f"Encontrado archivo con timestamp: {os.path.basename(latest_file)}")
        return latest_file
    
    # FALLBACK LEGACY: Buscar archivo sin timestamp (BASE.EXT)
    if legacy_fallback:
        legacy_pattern = os.path.join(directory, f"{base_name}{extension}")
        if os.path.exists(legacy_pattern):
            log.info(f"Usando archivo legacy (sin timestamp): {base_name}{extension}")
            return legacy_pattern
    
    log.warning(f"No se encontró ningún archivo para: {base_name}")
    return None


def find_all_timestamped_files(
    directory: str, 
    base_name: str, 
    extension: str = ".xlsx"
) -> List[str]:
    """
    Busca todos los archivos con timestamp para un nombre base dado.
    
    Útil para auditing o para mostrar todas las exportaciones disponibles
    de un archivo específico.
    
    Args:
        directory: Directorio donde buscar
        base_name: Nombre base del archivo
        extension: Extensión del archivo
    
    Returns:
        Lista de rutas completas, ordenadas de más reciente a más antigua
    """
    directory = os.path.normpath(directory)
    base_name = base_name.strip()
    
    timestamp_pattern = f"{base_name}__????????_??????{extension}"
    full_pattern = os.path.join(directory, timestamp_pattern)
    
    matching_files = glob.glob(full_pattern)
    matching_files.sort(reverse=True)
    
    return matching_files


def find_files_with_prefix(
    directory: str, 
    prefix: str, 
    extension: str = ".xlsx"
) -> List[str]:
    """
    Busca todos los archivos que empiecen por un prefijo dado.
    
    Útil para archivos con partes variables como:
    - SPA_Ventas_semana_42__20260205_210037.xlsx (donde 42 es la semana)
    - SPA_Stock__20260205_210037.xlsx
    
    Args:
        directory: Directorio donde buscar
        prefix: Prefijo del archivo
        extension: Extensión del archivo
    
    Returns:
        Lista de archivos encontrados, ordenados de más reciente a más antiguo
    """
    directory = os.path.normpath(directory)
    pattern = os.path.join(directory, f"{prefix}*{extension}")
    files = glob.glob(pattern)
    files.sort(reverse=True)
    return files


def extract_timestamp(filename: str) -> Optional[str]:
    """
    Extrae el timestamp de un nombre de archivo con formato ERP.
    
    Args:
        filename: Nombre del archivo (con o sin ruta)
    
    Returns:
        Timestamp en formato "YYYYMMDD_HHMMSS" o None si no tiene el formato esperado
    
    Examples:
        >>> extract_timestamp("SPA_Ventas__20260205_210037.xlsx")
        '20260205_210037'
        
        >>> extract_timestamp("Stock.xlsx")
        None
    """
    basename = os.path.basename(filename)
    # Buscar patrón: __ seguido de 8 dígitos, _, 6 dígitos, y termina en .xlsx
    match = re.search(r'__(\d{8}_\d{6})\.xlsx$', basename)
    if match:
        return match.group(1)
    return None


# ============================================
# PRUEBAS DEL MÓDULO
# ============================================

if __name__ == "__main__":
    import sys
    
    # Configurar logging básico para pruebas
    logging.basicConfig(
        level=logging.INFO,
        format='%(levelname)s: %(message)s'
    )
    
    print("=" * 60)
    print("PRUEBAS DEL MÓDULO file_finder.py")
    print("=" * 60)
    
    test_dir = "/workspace/PROGRAMA_PEDIDOS/LISTADO_PEDIDO_COMPRAS/data/input"
    
    # === PRUEBA 1: Importación y uso básico ===
    print("\n[TEST 1] Uso básico de find_latest_file()")
    print("-" * 40)
    
    result = find_latest_file(test_dir, "Ventas")
    print(f"Resultado: {result}")
    assert result is not None, "Debería encontrar Ventas.xlsx"
    print("✓ OK")
    
    # === PRUEBA 2: Extracción de timestamp ===
    print("\n[TEST 2] Extracción de timestamp")
    print("-" * 40)
    
    ts = extract_timestamp("SPA_Stock__20260205_210037.xlsx")
    print(f"Timestamp extraído: {ts}")
    assert ts == "20260205_210037", "Timestamp debería ser 20260205_210037"
    print("✓ OK")
    
    ts_none = extract_timestamp("Stock.xlsx")
    print(f"Timestamp de archivo sin timestamp: {ts_none}")
    assert ts_none is None, "Debería ser None"
    print("✓ OK")
    
    # === PRUEBA 3: Crear archivos de prueba y verificar selección ===
    print("\n[TEST 3] Verificar selección del archivo más reciente")
    print("-" * 40)
    
    # Crear archivos de prueba temporales
    test_files = [
        f"{test_dir}/PRUEBA_Arch__20260205_090000.xlsx",
        f"{test_dir}/PRUEBA_Arch__20260205_150000.xlsx",
        f"{test_dir}/PRUEBA_Arch__20260205_210037.xlsx"
    ]
    
    for f in test_files:
        Path(f).touch()
    
    result = find_latest_file(test_dir, "PRUEBA_Arch")
    print(f"Seleccionado: {os.path.basename(result)}")
    assert "210037" in result, "Debería seleccionar el más reciente"
    print("✓ OK")
    
    # Limpiar
    for f in test_files:
        if os.path.exists(f):
            os.remove(f)
    
    # === PRUEBA 4: Búsqueda por prefijo ===
    print("\n[TEST 4] Búsqueda por prefijo (semana variable)")
    print("-" * 40)
    
    # Crear archivo de prueba
    test_file = f"{test_dir}/PRUEBA_Ventas_semana_42__20260205_210037.xlsx"
    Path(test_file).touch()
    
    files = find_files_with_prefix(test_dir, "PRUEBA_Ventas_semana_42")
    print(f"Encontrados: {[os.path.basename(f) for f in files]}")
    assert len(files) == 1, "Debería encontrar exactamente 1 archivo"
    print("✓ OK")
    
    # Limpiar
    if os.path.exists(test_file):
        os.remove(test_file)
    
    print("\n" + "=" * 60)
    print("TODAS LAS PRUEBAS PASARON CORRECTAMENTE")
    print("=" * 60)
    sys.exit(0)
