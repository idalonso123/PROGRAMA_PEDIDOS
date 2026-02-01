#!/usr/bin/env python3
"""
Módulo ForecastEngine - Motor de cálculo de pedidos

Este módulo implementa la lógica principal de cálculo de pedidos de compra,
incluyendo la metodología de escalado proporcional, aplicación de factores de
crecimiento, ajustes por festividades, y cálculo de stock mínimo dinámico.
Toda la metodología está basada en el script original de pedido_compras.py.

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-01-31
"""

import pandas as pd
import numpy as np
import logging
import re
from typing import Optional, Dict, List, Any, Tuple
from datetime import datetime, date

# Configuración del logger
logger = logging.getLogger(__name__)


class ForecastEngine:
    """
    Motor de cálculo para la generación de pedidos de compra.
    
    Esta clase implementa toda la lógica matemática para calcular las cantidades
    de pedido óptimas para cada artículo, basándose en las ventas históricas,
    los objetivos de venta, la clasificación ABC, y los factores de ajuste
    configurados en el sistema.
    
    Attributes:
        config (dict): Configuración del sistema
        parametros (dict): Parámetros de cálculo (crecimiento, stock mínimo, etc.)
        festivos (dict): Factores de incremento por festividades
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el ForecastEngine con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.parametros = config.get('parametros', {})
        self.festivos = config.get('festivos', {})
        self.secciones = config.get('secciones', {})
        self.pesos_categoria = self.parametros.get('pesos_categoria', {
            'A': 1.0,
            'B': 0.8,
            'C': 0.6,
            'D': 0.0
        })
        
        logger.info("ForecastEngine inicializado correctamente")
    
    def calcular_factor_compra(self, accion_texto: Any) -> float:
        """
        Analiza la 'Acción Sugerida' y devuelve el factor de compra.
        
        El factor determina si se debe mantener, aumentar, reducir o eliminar
        las compras de un artículo según la estrategia comercial definida.
        
        Args:
            accion_texto (Any): Texto de la acción sugerida desde ABC+D
        
        Returns:
            float: Factor de compra (0=eliminar, 1=mantener, >1=aumentar)
        """
        if pd.isna(accion_texto):
            return 1.0
        
        accion = str(accion_texto).lower()
        
        # Eliminar acentos para comparación
        accion_normalizada = self._normalizar(accion)
        
        # ============================================
        # CASO ELIMINAR: Eliminar del catálogo
        # ============================================
        if 'eliminar del catalogo' in accion_normalizada:
            return 0.0
        
        # ============================================
        # CASOS DE REDUCCIÓN DE COMPRAS
        # ============================================
        reducciones = {
            'reducir compras 70%': 0.50,
            'reducir compras 50%': 0.50,
            'reducir compras 40%': 0.60,
            'reducir compras 35%': 0.65,
            'reducir compras 30%': 0.65,
            'reducir compras 25%': 0.75,
            'reducir compras 20%': 0.80,
            'reducir compras 15%': 0.85,
            'aplicar descuento 20%': 0.80,
            'implementar promocion del 15%': 0.85,
        }
        
        for patron, factor in reducciones.items():
            if patron in accion_normalizada:
                return factor
        
        # Aplicar descuento genérico (extraer porcentaje)
        if 'aplicar descuento' in accion_normalizada:
            match = re.search(r'aplicar descuento\s*(\d+[.,]?\d*)%', accion_normalizada)
            if match:
                porcentaje = float(match.group(1).replace(',', '.'))
                return 1.0 - (porcentaje / 100.0)
        
        # ============================================
        # CASOS DE MANTENER COMPRAS (0% cambio)
        # ============================================
        mantener_patrones = [
            'mantener el nivel de compras actual',
            'mantener nivel de compras',
            'mantener nivel de compras anterior'
        ]
        
        for patron in mantener_patrones:
            if patron in accion_normalizada:
                return 1.0
        
        # ============================================
        # CASOS DE AUMENTO DE COMPRAS
        # ============================================
        aumentos = {
            'aumentar compras 50%': 1.50,
            'aumentar compras 40%': 1.40,
            'incrementar compras 30%': 1.30,
            'aumentar compras 30%': 1.30,
            'aumentar compras 25%': 1.25,
            'incrementar compras 20%': 1.20,
            'aumentar compras 15%': 1.15,
        }
        
        for patron, factor in aumentos.items():
            if patron in accion_normalizada:
                return factor
        
        # ============================================
        # CASO POR DEFECTO: Mantener compras
        # ============================================
        return 1.0
    
    def _normalizar(self, texto: str) -> str:
        """
        Normaliza un texto eliminando acentos y convirtiendo a minúsculas.
        
        Args:
            texto (str): Texto a normalizar
        
        Returns:
            str: Texto normalizado
        """
        import unicodedata
        texto = texto.lower()
        texto = unicodedata.normalize('NFD', texto)
        texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
        return texto.strip()
    
    def obtener_numero_semana(self, fecha: datetime) -> int:
        """
        Obtiene el número de semana ISO para una fecha.
        
        Args:
            fecha (datetime): Fecha a evaluar
        
        Returns:
            int: Número de semana ISO (1-53)
        """
        return fecha.isocalendar()[1]
    
    def calcular_fechas_semana(self, semana: int, año: int = 2026) -> Tuple[str, str]:
        """
        Calcula las fechas de inicio y fin de una semana específica.
        
        Args:
            semana (int): Número de semana ISO
            año (int): Año de la semana
        
        Returns:
            Tuple[str, str]: (fecha_lunes, fecha_domingo) en formato YYYY-MM-DD
        """
        # Obtener el jueves de la semana (día central de la semana ISO)
        fecha = date(año, 1, 4)  # 4 de enero siempre está en la semana 1 del año ISO
        delta = datetime.timedelta(weeks=semana - 1)
        fecha = fecha + delta
        
        # Obtener el lunes de esa semana
        dias_lunes = fecha.weekday()
        lunes = fecha - datetime.timedelta(days=dias_lunes)
        
        # El fin de semana es el domingo
        domingo = lunes + datetime.timedelta(days=6)
        
        return lunes.strftime('%Y-%m-%d'), domingo.strftime('%Y-%m-%d')
    
    def obtener_objetivo_semana(self, seccion: str, semana: int) -> float:
        """
        Obtiene el objetivo de venta para una sección y semana específicas.
        
        Args:
            seccion (str): Nombre de la sección
            semana (int): Número de semana
        
        Returns:
            float: Objetivo de venta en euros
        """
        seccion_config = self.secciones.get(seccion, {})
        objetivos = seccion_config.get('objetivos_semanales', {})
        return objetivos.get(str(semana), 0.0)
    
    def calcular_pedido_semana(self, semana: int, datos_semana: pd.DataFrame,
                                abc_df: pd.DataFrame, costes_df: pd.DataFrame,
                                seccion: str) -> pd.DataFrame:
        """
        Calcula el pedido para una semana específica con la metodología correcta.
        
        La metodología implementada es:
        1. Aplicar filtros ABC+D a las unidades del año pasado
        2. Calcular ventas preliminares (unidades × PVP)
        3. Escalar las unidades para que las VENTAS ESCALADAS = OBJETIVO HISTÓRICO
        4. Aplicar crecimiento (+5%) sobre el OBJETIVO
        5. Aplicar festivo (+X%) sobre el OBJETIVO
        
        Args:
            semana (int): Número de semana a procesar
            datos_semana (pd.DataFrame): Datos de ventas históricas de la semana
            abc_df (pd.DataFrame): Datos de clasificación ABC
            costes_df (pd.DataFrame): Datos de costes y precios
            seccion (str): Nombre de la sección
        
        Returns:
            pd.DataFrame: DataFrame con los artículos y cantidades a pedir
        """
        if len(datos_semana) == 0:
            logger.warning(f"No hay datos para la semana {semana}")
            return pd.DataFrame()
        
        logger.info(f"Calculando pedido para semana {semana} ({len(datos_semana)} registros)")
        
        # Agrupar por artículo
        ventas_articulo = datos_semana.groupby(['Codigo', 'Nombre', 'Talla', 'Color']).agg({
            'Unidades': 'sum',
            'Importe': 'sum'
        }).reset_index()
        
        ventas_articulo.columns = ['Codigo', 'Nombre', 'Talla', 'Color', 'Unidades_Base', 'Importe_Base']
        
        # PASO 1: Aplicar lógica individualizada basada en "Acción Sugerida"
        pedidos = []
        
        for idx, row in ventas_articulo.iterrows():
            # Buscar información del artículo en ABC y costes
            info_articulo = self._buscar_info_articulo(
                row['Codigo'], row['Nombre'], row['Talla'], row['Color'],
                abc_df, costes_df
            )
            
            unidades_base = row['Unidades_Base']
            
            # Calcular factor de compra individualizado
            factor_compra = self.calcular_factor_compra(info_articulo['accion_raw'])
            
            # Aplicar el factor a las unidades base
            unidades_abc = unidades_base * factor_compra
            
            # Determinar texto de acción aplicada para referencia
            if factor_compra == 0:
                accion_aplicada = 'ELIMINAR'
            elif factor_compra < 1:
                accion_aplicada = f'REDUCIR {int((1-factor_compra)*100)}%'
            elif factor_compra > 1:
                accion_aplicada = f'AUMENTAR {int((factor_compra-1)*100)}%'
            else:
                accion_aplicada = 'MANTENER'
            
            pedidos.append({
                'Codigo_Articulo': row['Codigo'],
                'Nombre_Articulo': row['Nombre'],
                'Talla': row['Talla'],
                'Color': row['Color'],
                'Seccion': seccion,
                'Unidades_Base': unidades_base,
                'Unidades_ABC': unidades_abc,
                'PVP': info_articulo['pvp'],
                'Coste_Pedido': info_articulo['coste'],
                'Proveedor': info_articulo['proveedor'],
                'Categoria': info_articulo['categoria'],
                'Accion_Aplicada': accion_aplicada,
                'Peso_Categoria': self.pesos_categoria.get(info_articulo['categoria'], 0)
            })
        
        pedidos_df = pd.DataFrame(pedidos)
        
        # Calcular ventas preliminares
        pedidos_df['Ventas_Preliminares'] = pedidos_df['Unidades_ABC'] * pedidos_df['PVP']
        ventas_actuales = pedidos_df['Ventas_Preliminares'].sum()
        
        # Obtener objetivo de la semana
        objetivo_semana = self.obtener_objetivo_semana(seccion, semana)
        
        logger.info(f"  Objetivo: {objetivo_semana}€, Actual (ABC): {ventas_actuales:.2f}€")
        
        # Calcular factor de crecimiento y festivo
        crecimiento = self.parametros.get('objetivo_crecimiento', 0.05)
        festivo = self.festivos.get(str(semana), self.festivos.get(semana, 0.0))
        factor_total = (1 + crecimiento) * (1 + festivo)
        
        logger.info(f"  Factor crecimiento: {crecimiento}, Factor festivo: {festivo}")
        logger.info(f"  Factor total: {factor_total:.4f}")
        
        # Calcular factor de escalado
        if ventas_actuales > 0 and objetivo_semana > 0:
            factor_escalado = (objetivo_semana * factor_total) / ventas_actuales
        else:
            factor_escalado = 1.0
        
        logger.info(f"  Factor escalado: {factor_escalado:.4f}")
        
        # Calcular las unidades escaladas
        pedidos_df['Unidades_Escaladas'] = pedidos_df['Unidades_ABC'] * factor_escalado
        
        # Usar np.ceil para calcular unidades con redondeo hacia arriba
        pedidos_df['Unidades_Finales'] = pedidos_df['Unidades_Escaladas'].apply(
            lambda x: int(np.ceil(x)) if x > 0 else 0
        )
        
        # Calcular ventas preliminares con las unidades ceiling
        pedidos_df['Ventas_Preliminares'] = pedidos_df['Unidades_Finales'] * pedidos_df['PVP']
        ventas_preliminares = pedidos_df['Ventas_Preliminares'].sum()
        
        # El objetivo final es objetivo × factor_total
        objetivo_final = objetivo_semana * factor_total
        
        # Calcular el delta
        delta = ventas_preliminares - objetivo_final
        
        logger.info(f"  Ventas preliminares (con ceil): {ventas_preliminares:.2f}€")
        logger.info(f"  Objetivo final: {objetivo_final:.2f}€")
        logger.info(f"  Delta: {delta:.2f}€")
        
        # Si hay exceso (delta > 0), reducir unidades de los artículos con menor pvp
        if delta > 0:
            pedidos_df = pedidos_df.sort_values('PVP', ascending=True)
            
            reduction_needed = delta
            for idx in pedidos_df.index:
                if reduction_needed <= 0:
                    break
                current_units = pedidos_df.at[idx, 'Unidades_Finales']
                if current_units > 0:
                    pvp = pedidos_df.at[idx, 'PVP']
                    if pvp <= reduction_needed:
                        pedidos_df.at[idx, 'Unidades_Finales'] = current_units - 1
                        reduction_needed -= pvp
            
            pedidos_df['Ventas_Objetivo'] = (pedidos_df['Unidades_Finales'] * pedidos_df['PVP']).round(2)
        else:
            pedidos_df['Ventas_Objetivo'] = (pedidos_df['Unidades_Finales'] * pedidos_df['PVP']).round(2)
        
        # Calcular Beneficio Objetivo
        pedidos_df['Beneficio_Objetivo'] = (
            pedidos_df['Ventas_Objetivo'] - 
            (pedidos_df['Unidades_Finales'] * pedidos_df['Coste_Pedido'])
        ).round(2)
        
        ventas_finales = pedidos_df['Ventas_Objetivo'].sum()
        logger.info(f"  Ventas finales: {ventas_finales:.2f}€")
        
        return pedidos_df
    
    def aplicar_stock_minimo(self, pedidos_df: pd.DataFrame, semana: int,
                              stock_acumulado_dict: Dict[str, int]) -> Tuple[pd.DataFrame, Dict[str, int], Dict[str, int]]:
        """
        Aplica el cálculo de stock mínimo dinámico POR ARTÍCULO.
        
        Args:
            pedidos_df (pd.DataFrame): DataFrame con los pedidos calculados
            semana (int): Número de semana actual
            stock_acumulado_dict (Dict[str, int]): Stock acumulado por artículo
        
        Returns:
            Tuple: (pedidos_actualizados, nuevo_stock_acumulado, ajustes_articulo)
        """
        if len(pedidos_df) == 0:
            return pedidos_df, {}, {}
        
        stock_minimo_porcentaje = self.parametros.get('stock_minimo_porcentaje', 0.30)
        
        nuevo_stock_acumulado = {}
        ajustes_articulo = {}
        
        for idx, row in pedidos_df.iterrows():
            codigo = row['Codigo_Articulo']
            talla = row['Talla']
            color = row['Color']
            
            clave_articulo = f"{codigo}|{talla}|{color}"
            
            # Calcular stock mínimo individual basado en unidades finales
            stock_minimo = int(np.ceil(row['Unidades_Finales'] * stock_minimo_porcentaje))
            
            stock_acumulado_anterior = stock_acumulado_dict.get(clave_articulo, 0)
            diferencia_stock = stock_minimo - stock_acumulado_anterior
            
            unidades_pedido = row['Unidades_Finales'] + diferencia_stock
            
            pedidos_df.at[idx, 'Stock_Minimo_Objetivo'] = stock_minimo
            pedidos_df.at[idx, 'Unidades_Pedido'] = unidades_pedido
            pedidos_df.at[idx, 'Diferencia_Stock'] = diferencia_stock
            
            nuevo_stock_acumulado[clave_articulo] = stock_minimo
            ajustes_articulo[clave_articulo] = diferencia_stock
        
        return pedidos_df, nuevo_stock_acumulado, ajustes_articulo
    
    def _buscar_info_articulo(self, codigo: Any, nombre: Any, talla: Any, color: Any,
                               abc_df: pd.DataFrame, coste_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Busca la información del artículo en ABC+D y Coste.
        
        Args:
            codigo (Any): Código del artículo
            nombre (Any): Nombre del artículo
            talla (Any): Talla del artículo
            color (Any): Color del artículo
            abc_df (pd.DataFrame): DataFrame de clasificación ABC
            coste_df (pd.DataFrame): DataFrame de costes
        
        Returns:
            Dict: Información del artículo encontrada
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
                col_norm = col.lower().replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
                if 'nombre' in col_norm and 'proveedor' in col_norm:
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
        
        return {
            'accion_raw': accion_raw,
            'categoria': categoria,
            'descuento_sugerido': descuento_sugerido,
            'pvp': pvp,
            'coste': coste,
            'proveedor': proveedor
        }
    
    def generar_resumen_pedido(self, pedidos_df: pd.DataFrame, semana: int,
                                datos_originales: pd.DataFrame) -> Dict[str, Any]:
        """
        Genera un resumen consolidado del pedido para una semana.
        
        Args:
            pedidos_df (pd.DataFrame): DataFrame con los pedidos calculados
            semana (int): Número de semana
            datos_originales (pd.DataFrame): Datos originales de ventas
        
        Returns:
            Dict: Resumen con métricas del pedido
        """
        if len(pedidos_df) == 0:
            return {}
        
        # Filtrar artículos con Unidades_Pedido > 0
        pedidos_filtrados = pedidos_df[pedidos_df['Unidades_Pedido'] > 0]
        
        if len(pedidos_filtrados) == 0:
            return {}
        
        # Calcular ventas de la semana del año pasado
        datos_semana = datos_originales[datos_originales['Semana'] == semana]
        ventas_año_pasado = datos_semana['Importe'].sum() if len(datos_semana) > 0 else 0
        
        festivo = self.festivos.get(str(semana), self.festivos.get(semana, 0.0))
        crecimiento = self.parametros.get('objetivo_crecimiento', 0.05)
        factor_total = (1 + crecimiento) * (1 + festivo)
        
        objetivo = self.obtener_objetivo_semana(self._obtener_seccion_activa(), semana)
        objetivo_final = round(objetivo * factor_total, 2)
        
        crecimiento_unidades = round((objetivo_final / objetivo - 1) * 100, 1) if objetivo > 0 else 0
        
        return {
            'Semana': semana,
            'Vtas. semana año pasado': round(ventas_año_pasado, 2),
            'Objetivo_Semana': objetivo,
            'Obj. semana + % crec. anual': round(objetivo * (1 + crecimiento), 2),
            'Obj. semana + % crec. + Festivos': objetivo_final,
            '% Obj. crecim. + Festivos': crecimiento_unidades,
            'Total_Unidades': int(pedidos_filtrados['Unidades_Pedido'].sum()),
            'Total_Articulos': len(pedidos_filtrados),
            'Total_Importe': round(pedidos_filtrados['Ventas_Objetivo'].sum(), 2),
            'Alcance_Objetivo_%': round(pedidos_filtrados['Ventas_Objetivo'].sum() / objetivo * 100, 1) if objetivo > 0 else 0,
            'Articulos_A': len(pedidos_filtrados[pedidos_filtrados['Categoria'] == 'A']),
            'Articulos_B': len(pedidos_filtrados[pedidos_filtrados['Categoria'] == 'B']),
            'Articulos_C': len(pedidos_filtrados[pedidos_filtrados['Categoria'] == 'C']),
            'Incremento_Festivo_%': festivo * 100,
            'Stock_Minimo_%': self.parametros.get('stock_minimo_porcentaje', 0.30) * 100,
            'Stock_Minimo_Objetivo': int(pedidos_filtrados['Stock_Minimo_Objetivo'].sum())
        }
    
    def _obtener_seccion_activa(self) -> str:
        """
        Obtiene la sección activa de la configuración.
        
        Returns:
            str: Nombre de la sección activa
        """
        # Por defecto, devuelve la primera sección configurada
        secciones = self.config.get('secciones_activas', [])
        return secciones[0] if secciones else 'general'


# Funciones de utilidad para uso directo
def crear_forecast_engine(config: dict) -> ForecastEngine:
    """
    Crea una instancia del ForecastEngine.
    
    Args:
        config (dict): Configuración del sistema
    
    Returns:
        ForecastEngine: Instancia inicializada del motor de cálculo
    """
    return ForecastEngine(config)


if __name__ == "__main__":
    # Ejemplo de uso
    print("ForecastEngine - Motor de cálculo de pedidos")
    print("=" * 50)
    
    # Configurar logging básico
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'parametros': {
            'objetivo_crecimiento': 0.05,
            'stock_minimo_porcentaje': 0.30,
            'pesos_categoria': {'A': 1.0, 'B': 0.8, 'C': 0.6, 'D': 0.0}
        },
        'festivos': {'14': 0.25, '18': 0.00, '22': 0.00},
        'secciones_activas': ['vivero']
    }
    
    # Crear ForecastEngine
    engine = crear_forecast_engine(config_ejemplo)
    
    # Ejemplo de cálculo de factor de compra
    print("\nEjemplos de factor de compra:")
    print(f"  'Mantener nivel de compras': {engine.calcular_factor_compra('Mantener el nivel de compras actual')}")
    print(f"  'Aumentar compras 30%': {engine.calcular_factor_compra('Aumentar compras 30%')}")
    print(f"  'Reducir compras 20%': {engine.calcular_factor_compra('Reducir compras 20%')}")
    print(f"  'Eliminar del catálogo': {engine.calcular_factor_compra('Eliminar del catálogo')}")
    
    # Ejemplo de cálculo de fechas
    print(f"\nFechas de la semana 15 de 2026:")
    fecha_inicio, fecha_fin = engine.calcular_fechas_semana(15, 2026)
    print(f"  Del {fecha_inicio} al {fecha_fin}")
