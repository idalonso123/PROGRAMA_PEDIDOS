#!/usr/bin/env python3
"""
Módulo DataLoader - Carga y normalización de datos de entrada

Este módulo es responsable de toda la lectura y procesamiento de archivos Excel
que alimentan el sistema de generación de pedidos. Implementa funciones para
leer archivos de ventas, costes y clasificación ABC, con validación y limpieza
de datos automática.

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-01-31
"""

import pandas as pd
import numpy as np
import os
import glob
import logging
import unicodedata
from typing import Optional, Dict, List, Tuple, Any
from datetime import datetime

# Configuración del logger
logger = logging.getLogger(__name__)


class DataLoader:
    """
    Clase principal para la carga y normalización de datos de entrada.
    
    Proporciona métodos para leer archivos de ventas, costes y clasificación ABC,
    aplicando validaciones y transformaciones necesarias para garantizar la
    calidad de los datos utilizados en los cálculos de pedidos.
    
    Attributes:
        config (dict): Configuración del sistema cargada desde config.json
        rutas (dict): Rutas de archivos y directorios configurados
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el DataLoader con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.rutas = config.get('rutas', {})
        self.archivos = config.get('archivos_entrada', {})
        self.secciones = config.get('secciones_activas', [])
        self.codigos_mascotas = config.get('codigos_mascotas_vivo', [])
        
        logger.info("DataLoader inicializado correctamente")
    
    def normalizar_texto(self, texto: Any) -> str:
        """
        Normaliza un texto para comparaciones insensibles a mayúsculas y acentos.
        
        Esta función es fundamental para el matching de artículos entre diferentes
        fuentes de datos, ya que elimina las variaciones en la escritura que no
        afectan al significado.
        
        Args:
            texto (Any): Texto a normalizar (puede ser None, str o cualquier tipo)
        
        Returns:
            str: Texto normalizado, o cadena vacía si el input es None o no es string
        """
        if texto is None:
            return ''
        
        # Convertir a string si no lo es
        texto = str(texto)
        
        # Convertir a minúsculas
        texto = texto.lower()
        
        # Normalizar Unicode (eliminar acentos)
        texto = unicodedata.normalize('NFD', texto)
        texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
        
        # Eliminar espacios extra
        texto = texto.strip()
        
        return texto
    
    def contiene_texto(self, texto_buscar: str, texto_comparar: str) -> bool:
        """
        Compara si texto_buscar está contenido en texto_comparar.
        
        Args:
            texto_buscar (str): Texto que queremos encontrar
            texto_comparar (str): Texto donde buscamos
        
        Returns:
            bool: True si el texto está contenido (ignorando mayúsculas y acentos)
        """
        return self.normalizar_texto(texto_buscar) in self.normalizar_texto(texto_comparar)
    
    def texto_igual(self, texto1: Any, texto2: Any) -> bool:
        """
        Compara si dos textos son iguales.
        
        Args:
            texto1 (Any): Primer texto
            texto2 (Any): Segundo texto
        
        Returns:
            bool: True si son iguales (ignorando mayúsculas y acentos)
        """
        return self.normalizar_texto(texto1) == self.normalizar_texto(texto2)
    
    def determinar_seccion(self, codigo_articulo: Any) -> Optional[str]:
        """
        Determina la sección de un artículo según su código.
        
        Utiliza los rangos de códigos definidos en la configuración para clasificar
        cada artículo en su sección correspondiente. Esta clasificación es esencial
        para procesar los pedidos de forma correcta.
        
        Args:
            codigo_articulo (Any): Código del artículo (puede ser string o número)
        
        Returns:
            Optional[str]: Nombre de la sección o None si no se puede clasificar
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
        if len(codigo_str) < 10:
            return None
        
        # 1. Verificar códigos de mascotas vivos (primero, tienen prioridad)
        if codigo_str.startswith('2') and codigo_str[:4] in self.codigos_mascotas:
            return 'mascotas_vivo'
        
        # 2. Sección 2: Mascotas manufacturadas (empieza por 2 y no está en vivos)
        if codigo_str.startswith('2'):
            return 'mascotas_manufacturado'
        
        # 3. Sección 3: Tierra/Áridos (31 o 32)
        if codigo_str.startswith('31') or codigo_str.startswith('32'):
            return 'tierras_aridos'
        
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
    
    def obtener_directorio_entrada(self) -> str:
        """
        Obtiene el directorio de entrada configurado.
        
        Returns:
            str: Ruta del directorio de entrada
        """
        base = self.rutas.get('directorio_base', '.')
        entrada = self.rutas.get('directorio_entrada', './data/input')
        
        # Si es ruta relativa, combinar con base
        if not os.path.isabs(entrada):
            entrada = os.path.join(base, entrada)
        
        return entrada
    
    def obtener_directorio_salida(self) -> str:
        """
        Obtiene el directorio de salida configurado.
        
        Returns:
            str: Ruta del directorio de salida
        """
        base = self.rutas.get('directorio_base', '.')
        salida = self.rutas.get('directorio_salida', './data/output')
        
        # Si es ruta relativa, combinar con base
        if not os.path.isabs(salida):
            salida = os.path.join(base, salida)
        
        return salida
    
    def leer_excel(self, ruta_archivo: str, hoja: Optional[str] = None) -> Optional[pd.DataFrame]:
        """
        Lee un archivo Excel y devuelve un DataFrame.
        
        Args:
            ruta_archivo (str): Ruta del archivo Excel a leer
            hoja (Optional[str]): Nombre de la hoja a leer (None para todas)
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con los datos o None si hay error
        """
        try:
            if not os.path.exists(ruta_archivo):
                logger.error(f"Archivo no encontrado: {ruta_archivo}")
                return None
            
            logger.info(f"Leyendo archivo: {ruta_archivo}")
            
            if hoja:
                df = pd.read_excel(ruta_archivo, sheet_name=hoja)
            else:
                df = pd.read_excel(ruta_archivo, sheet_name=None)
            
            logger.info(f"Archivo leído exitosamente: {len(df) if isinstance(df, pd.DataFrame) else len(df)} hojas")
            return df
            
        except Exception as e:
            logger.error(f"Error al leer archivo {ruta_archivo}: {str(e)}")
            return None
    
    def leer_ventas(self) -> Optional[pd.DataFrame]:
        """
        Lee el archivo de ventas históricas.
        
        El archivo debe contener una hoja con datos de ventas por vendedor,
        incluyendo código de artículo, nombre, fecha, semana, unidades e importe.
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con las ventas procesadas o None
        """
        dir_entrada = self.obtener_directorio_entrada()
        nombre_archivo = self.archivos.get('ventas', 'SPA_ventas.xlsx')
        ruta_archivo = os.path.join(dir_entrada, nombre_archivo)
        
        df = self.leer_excel(ruta_archivo)
        
        if df is None:
            return None
        
        # Si devuelve diccionario de hojas, convertir a DataFrame
        if isinstance(df, dict):
            # Buscar hoja que contenga "ventas por vendedor"
            nombre_buscado = 'ventas por vendedor'
            nombre_encontrado = None
            
            for hoja in df.keys():
                if self.contiene_texto('ventas', hoja) and self.contiene_texto('vendedor', hoja):
                    nombre_encontrado = hoja
                    break
            
            if nombre_encontrado:
                df = df[nombre_encontrado]
                logger.info(f"Usando hoja: {nombre_encontrado}")
            else:
                # Usar la primera hoja disponible
                primera_hoja = list(df.keys())[0]
                df = df[primera_hoja]
                logger.warning(f"No se encontró hoja específica, usando: {primera_hoja}")
        else:
            df = df.copy()
        
        # Aplicar filtros de limpieza
        if 'Tipo registro' in df.columns:
            # Filtrar solo registros de tipo "Detalle"
            df = df[df['Tipo registro'].apply(lambda x: self.texto_igual(x, 'Detalle'))]
            logger.info(f"Filtrados {len(df)} registros de tipo 'Detalle'")
        
        if 'Artículo' in df.columns:
            # Eliminar filas con Artículo vacío
            registros_antes = len(df)
            df = df[df['Artículo'].notna()]
            registros_despues = len(df)
            if registros_antes != registros_despues:
                logger.info(f"Eliminados {registros_antes - registros_despues} registros con Artículo vacío")
        
        # Normalizar columnas de código y nombre
        if 'Artículo' in df.columns:
            df['Codigo'] = df['Artículo'].astype(str).str.strip()
        
        if 'Nombre artículo' in df.columns:
            df['Nombre'] = df['Nombre artículo'].astype(str).str.strip()
        
        logger.info(f"Ventas cargadas: {len(df)} registros")
        return df
    
    def leer_coste(self) -> Optional[pd.DataFrame]:
        """
        Lee el archivo de costes y precios.

        El archivo contiene información sobre PVP, costes de compra, proveedores
        y márgenes para cada artículo.

        Returns:
            Optional[pd.DataFrame]: DataFrame con los costes o None si hay error
        """
        dir_entrada = self.obtener_directorio_entrada()
        nombre_archivo = self.archivos.get('coste', 'SPA_coste.xlsx')
        ruta_archivo = os.path.join(dir_entrada, nombre_archivo)

        df = self.leer_excel(ruta_archivo)

        if df is None:
            return None

        # Si df es un diccionario (múltiples hojas), tomar la primera hoja
        if isinstance(df, dict):
            primera_hoja = list(df.keys())[0]
            df = df[primera_hoja]
            logger.debug(f"Usando primera hoja del archivo de costes: {primera_hoja}")

        df = df.copy()
        
        # Log de columnas disponibles para debugging
        logger.debug(f"Columnas disponibles en costes: {list(df.columns)}")

        # Determinar la columna de código del artículo
        columna_codigo = None
        for nombre_posible in ['Codigo', 'Artículo', 'Articulo', 'CODIGO', 'ARTÍCULO']:
            if nombre_posible in df.columns:
                columna_codigo = nombre_posible
                break
        
        if columna_codigo is None:
            logger.error(f"No se encontró columna de código en {nombre_archivo}. Columnas disponibles: {list(df.columns)}")
            return None
        
        logger.debug(f"Usando columna '{columna_codigo}' como código de artículo")

        # Normalizar columna de código
        df['Codigo'] = df[columna_codigo].astype(str).str.strip()

        # Verificar que existen las columnas necesarias
        for col in ['Talla', 'Color']:
            if col not in df.columns:
                logger.warning(f"Columna '{col}' no encontrada, inicializando con valores vacíos")
                df[col] = ''

        # Crear clave única para matching
        df['Clave'] = (df['Codigo'] + '|' +
                      df['Talla'].astype(str).str.strip() + '|' +
                      df['Color'].astype(str).str.strip())

        logger.info(f"Costes cargados: {len(df)} registros")
        return df
    
    def buscar_archivo_abc_seccion(self, seccion: str) -> Optional[str]:
        """
        Busca el archivo CLASIFICACION ABC+D específico para la sección.

        Busca en el directorio de entrada archivos que coincidan con el patrón
        'CLASIFICACION_ABC+D_' seguido del nombre de la sección, permitiendo
        también el nuevo formato con período y año (ej: CLASIFICACION_ABC+D_INTERIOR_P2_2025.xlsx).

        Args:
            seccion (str): Nombre de la sección a buscar

        Returns:
            Optional[str]: Ruta del archivo encontrado o None
        """
        dir_entrada = self.obtener_directorio_entrada()
        patron = self.archivos.get('clasificacion_abc', 'CLASIFICACION_ABC+D_*.xlsx')
        
        # Normalizar nombre de sección para búsqueda
        seccion_normalizada = self.normalizar_texto(seccion)
        
        # Buscar archivos con el nuevo formato que incluye período y año
        # Patrón: CLASIFICACION_ABC+D_{SECCION}_*.xlsx
        # Ejemplos: CLASIFICACION_ABC+D_INTERIOR_P1_2025.xlsx, CLASIFICACION_ABC+D_INTERIOR_P2_2026.xlsx
        nuevo_patron = f"CLASIFICACION_ABC+D_{seccion_normalizada}_*.xlsx"
        ruta_nueva = os.path.join(dir_entrada, nuevo_patron)
        archivos_nuevos = glob.glob(ruta_nueva)
        
        if archivos_nuevos:
            # Ordenar por fecha de modificación (más reciente primero)
            archivos_nuevos.sort(key=os.path.getmtime, reverse=True)
            logger.info(f"Archivo ABC encontrado (nuevo formato) para '{seccion}': {archivos_nuevos[0]}")
            return archivos_nuevos[0]
        
        # Si no encuentra el nuevo formato, buscar el formato antiguo
        # Patrón: CLASIFICACION_ABC+D_{SECCION}.xlsx (sin período)
        antiguo_patron = f"CLASIFICACION_ABC+D_{seccion_normalizada}.xlsx"
        ruta_antigua = os.path.join(dir_entrada, antiguo_patron)
        
        if os.path.exists(ruta_antigua):
            logger.info(f"Archivo ABC encontrado (formato antiguo) para '{seccion}': {ruta_antigua}")
            return ruta_antigua
        
        # Búsqueda amplia por sección (formato genérico)
        patron_generico = os.path.join(dir_entrada, f'*{seccion_normalizada}*.xlsx')
        archivos_genericos = glob.glob(patron_generico)
        
        if archivos_genericos:
            archivos_genericos.sort(key=os.path.getmtime, reverse=True)
            logger.warning(f"No se encontró archivo con patrón estándar para '{seccion}', usando: {archivos_genericos[0]}")
            return archivos_genericos[0]
        
        # Búsqueda genérica de cualquier archivo CLASIFICACION_ABC+D
        patron_catchall = os.path.join(dir_entrada, 'CLASIFICACION_ABC+D*.xlsx')
        archivos_catchall = glob.glob(patron_catchall)
        
        if archivos_catchall:
            archivos_catchall.sort(key=os.path.getmtime, reverse=True)
            logger.warning(f"No se encontró archivo específico para '{seccion}', usando: {archivos_catchall[0]}")
            return archivos_catchall[0]
        
        logger.error(f"No se encontró ningún archivo ABC+D para sección '{seccion}'")
        return None
    
    def leer_clasificacion_abc(self, seccion: str) -> Optional[pd.DataFrame]:
        """
        Lee el archivo de clasificación ABC para una sección específica.
        
        El archivo debe contener hojas separadas por categoría (A, B, C, D)
        con las acciones sugeridas y datos de rotación para cada artículo.
        
        Args:
            seccion (str): Nombre de la sección a procesar
        
        Returns:
            Optional[pd.DataFrame]: DataFrame con la clasificación ABC o None
        """
        ruta_archivo = self.buscar_archivo_abc_seccion(seccion)
        
        if ruta_archivo is None:
            return None
        
        df_dict = self.leer_excel(ruta_archivo)
        
        if df_dict is None or not isinstance(df_dict, dict):
            logger.error(f"Error al leer archivo ABC: {ruta_archivo}")
            return None
        
        # Combinar todas las hojas en un único DataFrame
        df_completo = []
        
        for hoja, df in df_dict.items():
            df = df.copy()
            
            # Extraer categoría del nombre de la hoja
            hoja_normalizada = self.normalizar_texto(hoja)
            
            if 'categoria a' in hoja_normalizada or 'categoría a' in hoja_normalizada:
                categoria = 'A'
            elif 'categoria b' in hoja_normalizada or 'categoría b' in hoja_normalizada:
                categoria = 'B'
            elif 'categoria c' in hoja_normalizada or 'categoría c' in hoja_normalizada:
                categoria = 'C'
            elif 'categoria d' in hoja_normalizada or 'categoría d' in hoja_normalizada:
                categoria = 'D'
            else:
                # Si no se puede determinar, asumir categoría C
                categoria = 'C'
                logger.warning(f"No se pudo determinar categoría para hoja '{hoja}', asignando 'C'")
            
            df['Categoria'] = categoria
            df_completo.append(df)
        
        # Concatenar todas las categorías
        df_resultado = pd.concat(df_completo, ignore_index=True)
        
        # Crear clave única para matching
        df_resultado['Clave'] = (df_resultado['Artículo'].astype(str).str.strip() + '|' + 
                                df_resultado['Nombre artículo'].astype(str).str.strip() + '|' + 
                                df_resultado['Talla'].astype(str).str.strip() + '|' + 
                                df_resultado['Color'].astype(str).str.strip())
        
        logger.info(f"Clasificación ABC cargada para '{seccion}': {len(df_resultado)} registros")
        return df_resultado
    
    def leer_datos_seccion(self, seccion: str) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """
        Lee todos los datos necesarios para procesar una sección.
        
        Args:
            seccion (str): Nombre de la sección a procesar
        
        Returns:
            Tuple: (abc_df, ventas_df, costes_df) o (None, None, None) si hay error
        """
        logger.info(f"=== LEYENDO DATOS PARA SECCIÓN: {seccion.upper()} ===")
        
        # DEBUG: Mostrar información de secciones activas
        logger.debug(f"[DEBUG] Secciones activas configuradas: {self.secciones}")
        
        # Leer clasificación ABC
        logger.debug(f"[DEBUG] Intentando leer clasificación ABC para sección: {seccion}")
        abc_df = self.leer_clasificacion_abc(seccion)
        
        logger.debug(f"[DEBUG] ABC leído: {len(abc_df) if abc_df is not None else 0} registros")
        if abc_df is not None and len(abc_df) > 0:
            logger.debug(f"[DEBUG] Columnas en ABC: {list(abc_df.columns)}")
        
        if abc_df is None:
            logger.error(f"No se pudo leer clasificación ABC para '{seccion}'")
            return None, None, None
        
        # Leer ventas
        logger.debug(f"[DEBUG] Intentando leer archivo de ventas...")
        ventas_df = self.leer_ventas()
        
        logger.debug(f"[DEBUG] Ventas leído: {len(ventas_df) if ventas_df is not None else 0} registros")
        if ventas_df is not None:
            logger.debug(f"[DEBUG] Columnas en Ventas: {list(ventas_df.columns)}")
        
        if ventas_df is None:
            logger.error("No se pudo leer archivo de ventas")
            return None, None, None
        
        # Filtrar ventas por sección
        logger.debug(f"[DEBUG] Filtrando ventas por sección: {seccion}")
        
        # DEBUG: Mostrar distribución de secciones antes de filtrar
        ventas_df['Seccion'] = ventas_df['Codigo'].apply(self.determinar_seccion)
        secciones_encontradas = ventas_df['Seccion'].value_counts()
        logger.debug(f"[DEBUG] Distribución de secciones antes de filtrar:")
        for sec, count in secciones_encontradas.items():
            logger.debug(f"  {sec}: {count} registros")
        
        registros_total = len(ventas_df)
        
        # DEBUG: Verificar si la sección solicitada existe
        if seccion not in secciones_encontradas.index:
            logger.warning(f"[DEBUG] La sección '{seccion}' NO se encontró en los datos")
            logger.debug(f"[DEBUG] Secciones disponibles: {list(secciones_encontradas.index)}")
        
        ventas_df = ventas_df[ventas_df['Seccion'] == seccion].copy()
        logger.debug(f"[DEBUG] Tras filtrar por '{seccion}': {len(ventas_df)} registros")
        logger.info(f"Filtrados {len(ventas_df)} registros de sección '{seccion}' de {registros_total} total")
        
        # Leer costes
        logger.debug(f"[DEBUG] Intentando leer archivo de costes...")
        costes_df = self.leer_coste()
        
        logger.debug(f"[DEBUG] Costes leído: {len(costes_df) if costes_df is not None else 0} registros")
        if costes_df is not None and len(costes_df) > 0:
            logger.debug(f"[DEBUG] Columnas en Costes: {list(costes_df.columns)}")
        
        if costes_df is None:
            logger.error("No se pudo leer archivo de costes")
            return None, None, None
        
        logger.info(f"=== DATOS LEÍDOS CORRECTAMENTE PARA {seccion.upper()} ===")
        logger.info(f"  ABC: {len(abc_df)} registros")
        logger.info(f"  Ventas: {len(ventas_df)} registros")
        logger.info(f"  Costes: {len(costes_df)} registros")
        
        return abc_df, ventas_df, costes_df
    
    def buscar_info_articulo(self, codigo: Any, nombre: Any, talla: Any, color: Any, 
                             abc_df: pd.DataFrame, coste_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Busca información de un artículo en ABC y en costes.
        
        Args:
            codigo (Any): Código del artículo
            nombre (Any): Nombre del artículo
            talla (Any): Talla del artículo
            color (Any): Color del artículo
            abc_df (pd.DataFrame): DataFrame de clasificación ABC
            coste_df (pd.DataFrame): DataFrame de costes
        
        Returns:
            Dict: Diccionario con la información encontrada del artículo
        """
        clave = f"{str(codigo).strip()}|{str(nombre).strip()}|{str(talla).strip()}|{str(color).strip()}"
        
        # Buscar en ABC+D
        accion_raw = None
        categoria = 'C'
        descuento_sugerido = 0
        
        match = abc_df[abc_df['Clave'] == clave]
        
        if len(match) > 0:
            accion_raw = match.iloc[0].get('Acción Sugerida')
            categoria = match.iloc[0].get('Categoria', 'C')
            descuento_sugerido = match.iloc[0].get('Descuento Sugerido (%)', 0)
        else:
            # Buscar solo por código
            match2 = abc_df[abc_df['Artículo'].astype(str) == str(codigo)]
            if len(match2) > 0:
                match3 = match2[(match2['Talla'].astype(str) == str(talla)) & 
                               (match2['Color'].fillna('').astype(str) == str(color))]
                if len(match3) > 0:
                    accion_raw = match3.iloc[0].get('Acción Sugerida')
                    categoria = match3.iloc[0].get('Categoria', 'C')
                    descuento_sugerido = match3.iloc[0].get('Descuento Sugerido (%)', 0)
                else:
                    accion_raw = match2.iloc[0].get('Acción Sugerida')
                    categoria = match2.iloc[0].get('Categoria', 'C')
                    descuento_sugerido = match2.iloc[0].get('Descuento Sugerido (%)', 0)
        
        # Buscar PVP, Coste y Proveedor
        clave_coste = f"{str(codigo).strip()}|{str(talla).strip()}|{str(color).strip()}"
        match_coste = coste_df[coste_df['Clave'] == clave_coste]
        
        pvp = 0
        coste = 0
        proveedor = ''
        
        if len(match_coste) > 0:
            pvp = match_coste.iloc[0].get('Tarifa10', 0)
            coste = match_coste.iloc[0].get('Coste', 0)
            
            # Buscar nombre del proveedor
            for col in match_coste.columns:
                if self.normalizar_texto(col) == 'nombre proveedor':
                    proveedor = match_coste.iloc[0][col]
                    break
            
            if pd.isna(proveedor):
                proveedor = ''
            
            # Calcular PVP por defecto si es 0
            if pvp == 0 or pd.isna(pvp):
                pvp = coste * 2.5
            
            # Calcular coste por defecto si es 0
            if coste == 0 or pd.isna(coste):
                coste = pvp / 2.5
        
        # BÚSQUEDA INDEPENDIENTE DE PROVEEDOR (igual que el script original):
        # Si no se encontró proveedor en la primera búsqueda, buscar solo por código
        # Esto es independiente del resto de columnas y no afecta a PVP ni Coste
        if proveedor == '' or pd.isna(proveedor):
            # Buscar todos los artículos con el mismo código
            match_proveedor = coste_df[coste_df['Codigo'].astype(str) == str(codigo).strip()]
            if len(match_proveedor) > 0:
                # Buscar la columna del proveedor
                columnas_normalizadas = [self.normalizar_texto(col) for col in match_proveedor.columns]
                nombre_proveedor_normalizado = self.normalizar_texto('Nombre proveedor')
                idx_proveedor = None
                for i, col_norm in enumerate(columnas_normalizadas):
                    if nombre_proveedor_normalizado == col_norm:
                        idx_proveedor = i
                        break
                
                # Buscar un registro que tenga proveedor válido
                for idx, row in match_proveedor.iterrows():
                    if idx_proveedor is not None:
                        prov = row.iloc[idx_proveedor]
                    else:
                        prov = ''
                    if pd.notna(prov) and prov != '':
                        proveedor = prov
                        break
        
        return {
            'accion_raw': accion_raw,
            'categoria': categoria,
            'descuento_sugerido': descuento_sugerido,
            'pvp': pvp,
            'coste': coste,
            'proveedor': proveedor
        }


# Funciones de utilidad para uso directo
def cargar_configuracion(ruta_config: str = 'config/config.json') -> Optional[dict]:
    """
    Carga la configuración desde un archivo JSON.
    
    Args:
        ruta_config (str): Ruta al archivo de configuración
    
    Returns:
        Optional[dict]: Configuración cargada o None si hay error
    """
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            config = pd.read_json(f)
        
        logger.info(f"Configuración cargada desde: {ruta_config}")
        return config
        
    except Exception as e:
        logger.error(f"Error al cargar configuración: {str(e)}")
        return None


if __name__ == "__main__":
    # Ejemplo de uso
    print("DataLoader - Módulo de carga de datos")
    print("=" * 50)
    
    # Configurar logging básico
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de inicialización
    config_ejemplo = {
        'rutas': {
            'directorio_base': '.',
            'directorio_entrada': './data/input',
            'directorio_salida': './data/output'
        },
        'archivos_entrada': {
            'ventas': 'SPA_ventas.xlsx',
            'coste': 'SPA_coste.xlsx'
        },
        'secciones_activas': ['vivero', 'interior'],
        'codigos_mascotas_vivo': ['2104', '2204']
    }
    
    loader = DataLoader(config_ejemplo)
    print("DataLoader inicializado y listo para usar.")
