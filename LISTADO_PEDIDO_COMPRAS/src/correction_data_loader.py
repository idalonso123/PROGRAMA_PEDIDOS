#!/usr/bin/env python3
"""
Módulo CorrectionDataLoader - Carga de datos de corrección para FASE 2

Este módulo extiende el DataLoader existente para leer el archivo de stock
actual que alimenta el sistema de corrección de pedidos:
- Stock_actual.xlsx: Inventario disponible al momento del cálculo

La integración de estos datos permite ajustar las proyecciones teóricas de la
FASE 1 contra la realidad operativa del almacén, corrigiendo el pedido generado
en función de la diferencia entre el stock real y el stock mínimo configurado.

Autor: Sistema de Pedidos Vivero V2 - FASE 2
Fecha: 2026-02-03
"""

import pandas as pd
import numpy as np
import os
import glob
import logging
import unicodedata
from typing import Optional, Dict, List, Tuple, Any
from datetime import datetime
from src.data_loader import DataLoader
from src.file_finder import find_latest_file

# Configuración del logger
logger = logging.getLogger(__name__)


class CorrectionDataLoader:
    """
    Clase para la carga y normalización de datos de corrección.
    
    Extiende el DataLoader base para manejar el archivo específico
    de stock actual de la FASE 2.
    
    Attributes:
        config (dict): Configuración del sistema
        rutas (dict): Rutas de archivos y directorios
        correction_files (dict): Rutas de archivos de corrección
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el CorrectionDataLoader con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.rutas = config.get('rutas', {})
        self.correction_files = config.get('archivos_correccion', {})
        
        # Usar el DataLoader base para funciones compartidas
        self.base_loader = DataLoader(config)
        
        logger.info("CorrectionDataLoader inicializado correctamente")
    
    def normalizar_texto(self, texto: Any) -> str:
        """
        Normaliza un texto para comparaciones insensibles a mayúsculas y acentos.
        
        Args:
            texto (Any): Texto a normalizar
        
        Returns:
            str: Texto normalizado
        """
        return self.base_loader.normalizar_texto(texto)
    
    def texto_igual(self, texto1: Any, texto2: Any) -> bool:
        """
        Compara si dos textos son iguales ignorando mayúsculas y acentos.
        
        Args:
            texto1 (Any): Primer texto
            texto2 (Any): Segundo texto
        
        Returns:
            bool: True si son iguales
        """
        return self.base_loader.normalizar_texto(texto1) == self.base_loader.normalizar_texto(texto2)
    
    def obtener_directorio_entrada(self) -> str:
        """
        Obtiene el directorio de entrada configurado.
        
        Returns:
            str: Ruta del directorio de entrada
        """
        base = self.rutas.get('directorio_base', '.')
        entrada = self.rutas.get('directorio_entrada', './data/input')
        
        if not os.path.isabs(entrada):
            entrada = os.path.join(base, entrada)
        
        return entrada
    
    def leer_excel(self, ruta_archivo: str, hoja: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Lee un archivo Excel y devuelve un DataFrame.
        
        Args:
            ruta_archivo (str): Ruta del archivo Excel
            hoja (Optional[str]): Nombre de la hoja (None para todas)
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con los datos o None si hay error
        """
        try:
            if not os.path.exists(ruta_archivo):
                logger.warning(f"Archivo de corrección no encontrado: {ruta_archivo}")
                return None
            
            logger.info(f"Leyendo archivo de corrección: {ruta_archivo}")
            
            if hoja:
                df = pd.read_excel(ruta_archivo, sheet_name=hoja)
            else:
                df = pd.read_excel(ruta_archivo, sheet_name=None)
            
            logger.info(f"Archivo leído exitosamente: {len(df) if isinstance(df, pd.DataFrame) else len(df)} hojas")
            return df
            
        except Exception as e:
            logger.error(f"Error al leer archivo {ruta_archivo}: {str(e)}")
            return None
    
    def buscar_archivo_correccion(self, nombre_archivo: str) -> Optional[str]:
        """
        Busca un archivo de corrección en el directorio de entrada.
        
        Args:
            nombre_archivo (str): Nombre del archivo a buscar
        
        Returns:
            Optional[str]: Ruta completa del archivo o None si no existe
        """
        dir_entrada = self.obtener_directorio_entrada()
        ruta_archivo = os.path.join(dir_entrada, nombre_archivo)
        
        if os.path.exists(ruta_archivo):
            logger.info(f"Archivo de corrección encontrado: {ruta_archivo}")
            return ruta_archivo
        
        # Intentar búsqueda con wildcards
        patron_buscar = os.path.join(dir_entrada, f"*{nombre_archivo}*")
        archivos_encontrados = glob.glob(patron_buscar)
        
        if archivos_encontrados:
            logger.info(f"Archivo encontrado (búsqueda amplia): {archivos_encontrados[0]}")
            return archivos_encontrados[0]
        
        logger.warning(f"Archivo de corrección no encontrado: {nombre_archivo}")
        return None
    
    def leer_stock_actual(self, semana: Optional[int] = None) -> Optional[pd.DataFrame]:
        """
        Lee el archivo de stock actual (SPA_Stock_actual__YYYYMMDD_HHMMSS.xlsx).
        
        Este archivo contiene el inventario disponible al momento del cálculo,
        incluyendo código de artículo, nombre, talla, color, unidades en stock,
        fecha del último movimiento y antigüedad del stock.
        
        Utiliza búsqueda flexible para encontrar archivos con timestamp del ERP
        o fallback a archivo legacy.
        
        Args:
            semana (Optional[int]): Número de semana (no usado actualmente, para futuro)
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con el stock actual o None si hay error
        """
        dir_entrada = self.obtener_directorio_entrada()
        
        # Usar búsqueda flexible de archivos
        # El config.json debe tener: "stock_actual": "SPA_Stock_actual" (sin extensión)
        base_name = self.correction_files.get('stock_actual', 'Stock_actual').replace('.xlsx', '')
        ruta = find_latest_file(dir_entrada, base_name, logger_instance=logger)
        
        if ruta is None:
            logger.error(f"No se encontró archivo de stock actual: {base_name}")
            return None
        
        logger.info(f"Leyendo archivo de stock: {os.path.basename(ruta)}")
        df = self.leer_excel(ruta)
        
        if df is None:
            return None
        
        # Si devuelve diccionario (múltiples hojas), tomar la primera
        if isinstance(df, dict):
            primera_hoja = list(df.keys())[0]
            df = df[primera_hoja]
            logger.debug(f"Usando hoja: {primera_hoja}")
        
        df = df.copy()
        
        # Normalizar columna de código de artículo
        self._normalizar_columnas_stock(df)
        
        # Validar que tenemos columna de stock
        if 'Stock_Fisico' not in df.columns and 'Stock' not in df.columns:
            # Buscar columna que contenga 'stock'
            for col in df.columns:
                if 'stock' in self.normalizar_texto(col):
                    df.rename(columns={col: 'Stock_Fisico'}, inplace=True)
                    break
        
        logger.info(f"Stock actual cargado: {len(df)} registros")
        return df
    
    def _normalizar_columnas_stock(self, df: pd.DataFrame) -> None:
        """
        Normaliza las columnas del DataFrame de stock.
        
        Args:
            df (pd.DataFrame): DataFrame a normalizar
        """
        # Renombrar columnas comunes a nombres estándar
        mapeo_columnas = {}
        
        for col in df.columns:
            col_norm = self.normalizar_texto(col)
            
            if 'articulo' in col_norm and 'codigo' in col_norm:
                mapeo_columnas[col] = 'Codigo_Articulo'
            elif 'codigo' in col_norm:
                mapeo_columnas[col] = 'Codigo_Articulo'
            elif 'nombre' in col_norm and 'articulo' in col_norm:
                mapeo_columnas[col] = 'Nombre_Articulo'
            elif col_norm == 'nombre':
                mapeo_columnas[col] = 'Nombre_Articulo'
            elif 'stock' in col_norm and ('fisico' in col_norm or 'actual' in col_norm or col_norm == 'stock'):
                mapeo_columnas[col] = 'Stock_Fisico'
            elif 'unidades' in col_norm and 'stock' in col_norm:
                mapeo_columnas[col] = 'Stock_Fisico'
            elif 'talla' in col_norm:
                mapeo_columnas[col] = 'Talla'
            elif 'color' in col_norm:
                mapeo_columnas[col] = 'Color'
            elif 'fecha' in col_norm and 'ultimo' in col_norm:
                mapeo_columnas[col] = 'Fecha_Ultimo_Movimiento'
            elif 'antiguedad' in col_norm:
                mapeo_columnas[col] = 'Antiguedad_Stock'
        
        if mapeo_columnas:
            df.rename(columns=mapeo_columnas, inplace=True)
            logger.debug(f"Columnas renombradas en stock: {list(mapeo_columnas.values())}")
    
    def cargar_datos_correccion(self, semana: Optional[int] = None) -> Dict[str, Optional[pd.DataFrame]]:
        """
        Carga los datos de corrección para una semana específica.
        
        Args:
            semana (Optional[int]): Número de semana para la que cargar datos
        
        Returns:
            Dict[str, Optional[pd.DataFrame]]: Diccionario con:
                - 'stock': DataFrame de stock actual
        """
        logger.info("=" * 60)
        logger.info(f"CARGANDO DATOS DE CORRECCIÓN PARA SEMANA {semana if semana else 'actual'}")
        logger.info("=" * 60)
        
        datos = {
            'stock': None
        }
        
        # Cargar stock actual
        datos['stock'] = self.leer_stock_actual(semana)
        if datos['stock'] is not None:
            logger.info(f"  Stock: {len(datos['stock'])} registros")
        else:
            logger.warning("No se encontraron datos de stock para la corrección")
        
        return datos
    
    def merge_con_pedido_teorico(
        self, 
        pedido_teorico: pd.DataFrame, 
        datos_correccion: Dict[str, Optional[pd.DataFrame]],
        clave_cols: List[str] = ['Codigo_Articulo', 'Talla', 'Color']
    ) -> pd.DataFrame:
        """
        Fusiona los datos de corrección con el pedido teórico de la FASE 1.
        
        Realiza un LEFT JOIN para mantener todos los artículos del pedido
        teórico, añadiendo las columnas de datos de stock real.
        
        Args:
            pedido_teorico (pd.DataFrame): DataFrame del pedido generado en FASE 1
            datos_correccion (Dict): Diccionario con datos de corrección
            clave_cols (List[str]): Columnas usadas como clave de unión
        
        Returns:
            pd.DataFrame: DataFrame fusionado con todos los datos
        """
        logger.info("Fusionando datos de corrección con pedido teórico...")
        
        df = pedido_teorico.copy()
        
        # Preparar claves de unión normalizadas
        df['_clave'] = (
            df.get('Codigo_Articulo', df.get('Código artículo', df.get('Codigo', ''))).astype(str) + '|' +
            df.get('Talla', '').astype(str) + '|' +
            df.get('Color', '').astype(str)
        )
        
        # Fusionar stock actual
        if datos_correccion['stock'] is not None:
            stock_df = datos_correccion['stock'].copy()
            stock_df['_clave'] = (
                stock_df.get('Codigo_Articulo', '').astype(str) + '|' +
                stock_df.get('Talla', '').astype(str) + '|' +
                stock_df.get('Color', '').astype(str)
            )
            
            # Seleccionar columnas relevantes
            cols_stock = ['_clave', 'Stock_Fisico']
            cols_disponibles = [c for c in cols_stock if c in stock_df.columns]
            stock_df = stock_df[cols_disponibles]
            
            # Agrupar por clave (si hay duplicados)
            stock_df = stock_df.groupby('_clave').agg({
                'Stock_Fisico': 'sum' if 'Stock_Fisico' in stock_df.columns else 'first'
            }).reset_index()
            
            df = df.merge(stock_df, on='_clave', how='left')
            
            # Rellenar NaN con 0
            if 'Stock_Fisico' in df.columns:
                df['Stock_Fisico'] = df['Stock_Fisico'].fillna(0)
        
        # Limpiar columna de clave temporal
        df.drop(columns=['_clave'], inplace=True, errors='ignore')
        
        # Añadir columna de stock faltante con valor por defecto
        if 'Stock_Fisico' not in df.columns:
            df['Stock_Fisico'] = 0
        
        logger.info(f"Fusión completada: {len(df)} registros")
        
        return df


# Función de utilidad para uso directo
def crear_correction_data_loader(config: dict) -> CorrectionDataLoader:
    """
    Crea una instancia del CorrectionDataLoader.
    
    Args:
        config (dict): Configuración del sistema
    
    Returns:
        CorrectionDataLoader: Instancia del loader de corrección
    """
    return CorrectionDataLoader(config)


if __name__ == "__main__":
    # Ejemplo de uso
    print("CorrectionDataLoader - Módulo de carga de datos de corrección FASE 2")
    print("=" * 60)
    
    # Configurar logging
    import logging
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración mínima
    config_ejemplo = {
        'rutas': {
            'directorio_base': '.',
            'directorio_entrada': './data/input'
        },
        'archivos_correccion': {
            'stock_actual': 'Stock_actual.xlsx'
        }
    }
    
    loader = CorrectionDataLoader(config_ejemplo)
    print("CorrectionDataLoader inicializado y listo para usar.")
