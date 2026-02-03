#!/usr/bin/env python3
"""
Módulo CorrectionEngine - Motor de corrección de pedidos FASE 2

Este módulo implementa la lógica de corrección basada en la matriz de escenarios
documentada en el archivo OPCIONES.docx. El sistema ajusta las proyecciones
teóricas de la FASE 1 contra la realidad operativa del almacén.

FÓRMULA PRINCIPAL (satisface todos los 54 escenarios):
    Pedido_Final = max(0, Pedido_Generado + (Stock_Mínimo - Stock_Real))

Donde:
    - Pedido_Generado: Resultado del algoritmo de predicción (FASE 1)
    - Stock_Mínimo: Definido por la clasificación ABC (Safety Stock)
    - Stock_Real: Dato físico proveniente del archivo Stock_actual.xlsx
    - max(0, ...): Función de corrección para evitar pedidos negativos

Autor: Sistema de Pedidos Vivero V2 - FASE 2
Fecha: 2026-02-03
"""

import pandas as pd
import numpy as np
import logging
from typing import Optional, Dict, List, Tuple, Any
from dataclasses import dataclass
from enum import Enum

# Configuración del logger
logger = logging.getLogger(__name__)


class CategoriaABC(Enum):
    """Enumeración de categorías ABC."""
    A = "A"
    B = "B"
    C = "C"
    D = "D"


@dataclass
class ConfiguracionCorreccion:
    """
    Configuración para el motor de corrección.
    
    Attributes:
        politica_stock_minimo (Dict): Stock mínimo por categoría ABC
        umbral_alerta_stock (int): Umbral para generar alertas de stock bajo
        permitir_pedidos_negativos (bool): Si True, permite valores negativos
        aplicar_tendencia (bool): Si True, ajusta por tendencia de ventas
        semanas_tendencia (int): Número de semanas para calcular tendencia
    """
    politica_stock_minimo: Dict[str, float] = None
    umbral_alerta_stock: int = 0
    permitir_pedidos_negativos: bool = False
    aplicar_tendencia: bool = False
    semanas_tendencia: int = 4
    
    def __post_init__(self):
        if self.politica_stock_minimo is None:
            self.politica_stock_minimo = {
                'A': 1.5,  # 1.5 semanas de cobertura para categoría A
                'B': 1.0,  # 1 semana de cobertura para categoría B
                'C': 0.5,  # 0.5 semanas de cobertura para categoría C
                'D': 0.0   # Sin stock mínimo para categoría D
            }


class CorrectionEngine:
    """
    Motor de corrección de pedidos basado en la matriz de escenarios OPCIONES.
    
    Este motor aplica la lógica de corrección para ajustar los pedidos teóricos
    de la FASE 1 contra la realidad operativa (stock real, ventas reales,
    compras reales).
    
    La implementación se basa en la fórmula general extraída del análisis
    de los 54 escenarios del documento OPCIONES:
    
        Pedido_Final = max(0, Pedido_Generado + Diferencia_Stock)
        
    Donde Diferencia_Stock = Stock_Mínimo - Stock_Real
    
    Attributes:
        config (ConfiguracionCorreccion): Configuración del motor
        abc_config (Dict): Configuración ABC del sistema
    """
    
    def __init__(
        self, 
        config_abc: Optional[Dict[str, Any]] = None,
        configuracion: Optional[ConfiguracionCorreccion] = None
    ):
        """
        Inicializa el CorrectionEngine.
        
        Args:
            config_abc (Optional[Dict]): Configuración ABC del sistema
            configuracion (Optional[ConfiguracionCorreccion]): Configuración específica
        """
        self.config = configuracion or ConfiguracionCorreccion()
        self.abc_config = config_abc or {}
        
        logger.info("CorrectionEngine inicializado correctamente")
        logger.debug(f"Política de stock mínimo: {self.config.politica_stock_minimo}")
    
    def calcular_diferencia_stock(
        self, 
        stock_minimo: float, 
        stock_real: float
    ) -> float:
        """
        Calcula la diferencia entre el stock mínimo objetivo y el stock real.
        
        Args:
            stock_minimo (float): Stock mínimo objetivo
            stock_real (float): Stock real actual
        
        Returns:
            float: Diferencia (puede ser positiva o negativa)
        """
        return stock_minimo - stock_real
    
    def aplicar_formula_correccion(
        self, 
        pedido_generado: float, 
        stock_minimo: float, 
        stock_real: float
    ) -> float:
        """
        Aplica la fórmula principal de corrección.
        
        Esta es la fórmula unificada que satisface todos los escenarios
        de la matriz OPCIONES:
        
            Pedido_Final = max(0, Pedido_Generado + (Stock_Mínimo - Stock_Real))
        
        Args:
            pedido_generado (float): Cantidad sugerida por FASE 1
            stock_minimo (float): Stock mínimo objetivo del artículo
            stock_real (float): Stock real actual en almacén
        
        Returns:
            float: Pedido corregido (nunca negativo si permitir_pedidos_negativos=False)
        """
        diferencia = self.calcular_diferencia_stock(stock_minimo, stock_real)
        pedido_corregido = pedido_generado + diferencia
        
        if self.config.permitir_pedidos_negativos:
            return pedido_corregido
        else:
            return max(0, pedido_corregido)
    
    def obtener_stock_minimo(
        self, 
        categoria_abc: str, 
        ventas_promedio: float = 0,
        pedido_generado: float = 0
    ) -> float:
        """
        Obtiene el stock mínimo para un artículo según su categoría ABC.
        
        Args:
            categoria_abc (str): Categoría ABC del artículo (A, B, C, D)
            ventas_promedio (float): Ventas promedio semanales
            pedido_generado (float): Pedido generado en FASE 1
        
        Returns:
            float: Stock mínimo objetivo
        """
        # Si hay configuración ABC específica con porcentaje de stock mínimo
        if 'pesos_categoria' in self.abc_config:
            pesos = self.abc_config['pesos_categoria']
            if categoria_abc in pesos:
                peso = pesos[categoria_abc]
                return pedido_generado * peso * self.config.politica_stock_minimo.get(categoria_abc, 0)
        
        # Usar política de semanas de cobertura
        semanas_cobertura = self.config.politica_stock_minimo.get(
            categoria_abc.upper(), 
            self.config.politica_stock_minimo.get('C', 0.5)
        )
        
        # Calcular como semanas de cobertura de las ventas promedio
        if ventas_promedio > 0:
            return ventas_promedio * semanas_cobertura
        else:
            # Si no hay ventas promedio, usar el pedido generado como base
            return pedido_generado * semanas_cobertura
    
    def detectar_escenario(
        self, 
        pedido_generado: float,
        stock_minimo: float,
        stock_real: float,
        ventas_reales: float = 0,
        ventas_objetivo: float = 0,
        compras_reales: float = 0,
        compras_sugeridas: float = 0
    ) -> Dict[str, Any]:
        """
        Detecta el escenario actual basándose en las variables del artículo.
        
        Esta función identifica cuál de los 54 escenarios de la matriz OPCIONES
        aplica al artículo, basándose en las comparaciones entre las variables.
        
        Args:
            pedido_generado (float): Pedido sugerido por FASE 1
            stock_minimo (float): Stock mínimo objetivo
            stock_real (float): Stock real actual
            ventas_reales (float): Ventas reales de la semana
            ventas_objetivo (float): Ventas objetivo de la semana
            compras_reales (float): Compras recibidas en la semana
            compras_sugeridas (float): Compras que debían llegar según FASE 1
        
        Returns:
            Dict[str, Any]: Información del escenario detectado
        """
        escenario = {
            'codigo': None,
            'descripcion': '',
            'ventas_vs_objetivo': None,
            'compras_vs_sugerido': None,
            'stock_vs_minimo': None,
            'requiere_correccion': False,
            'tipo_correccion': ''
        }
        
        # Comparar ventas reales vs objetivo
        if ventas_reales > ventas_objetivo:
            escenario['ventas_vs_objetivo'] = 'SUPERIOR'
        elif ventas_reales < ventas_objetivo:
            escenario['ventas_vs_objetivo'] = 'INFERIOR'
        else:
            escenario['ventas_vs_objetivo'] = 'IGUAL'
        
        # Comparar compras reales vs sugerido
        if compras_reales > compras_sugeridas:
            escenario['compras_vs_sugerido'] = 'EXCESO'
        elif compras_reales < compras_sugeridas:
            escenario['compras_vs_sugerido'] = 'DEFECTO'
        else:
            escenario['compras_vs_sugerido'] = 'IGUAL'
        
        # Comparar stock real vs mínimo
        if stock_real > stock_minimo:
            escenario['stock_vs_minimo'] = 'EXCEDENTE'
        elif stock_real < stock_minimo:
            escenario['stock_vs_minimo'] = 'DEFICIT'
        else:
            escenario['stock_vs_minimo'] = 'OPTIMO'
        
        # Generar código de escenario
        escenario['codigo'] = (
            f"{escenario['ventas_vs_objetivo'][:3]}_"
            f"{escenario['compras_vs_sugerido'][:3]}_"
            f"{escenario['stock_vs_minimo'][:3]}"
        )
        
        # Determinar tipo de corrección
        if stock_real >= stock_minimo:
            escenario['requiere_correccion'] = stock_real > stock_minimo
            escenario['tipo_correccion'] = 'REDUCIR_EXCEDENTE' if stock_real > stock_minimo else 'MANTENER'
        else:
            escenario['requiere_correccion'] = True
            escenario['tipo_correccion'] = 'RECUPERAR_DEFICIT'
        
        # Generar descripción
        escenarios_descriptivos = {
            ('SUPERIOR', 'EXCESO', 'DEFICIT'): 
                'Ventas altas y exceso de compras generaron déficit de stock',
            ('SUPERIOR', 'EXCESO', 'OPTIMO'):
                'Ventas altas compensaron exceso de compras',
            ('SUPERIOR', 'EXCESO', 'EXCEDENTE'):
                'Exceso de compras con ventas altas pero aún hay excedente',
            ('SUPERIOR', 'IGUAL', 'DEFICIT'):
                'Ventas altas sin compras adicionales generaron déficit',
            ('SUPERIOR', 'IGUAL', 'OPTIMO'):
                'Ventas altas compensaron exactamente las compras',
            ('SUPERIOR', 'IGUAL', 'EXCEDENTE'):
                'Ventas altas pero no suficientes para compensar compras',
            ('SUPERIOR', 'DEFECTO', 'DEFICIT'):
                'Ventas altas con pocas compras: déficit crítico',
            ('SUPERIOR', 'DEFECTO', 'OPTIMO'):
                'Ventas altas pero compras justas mantienen stock óptimo',
            ('SUPERIOR', 'DEFECTO', 'EXCEDENTE'):
                '即使购买不足，高销量仍有剩余库存',
            ('IGUAL', 'EXCESO', 'DEFICIT'):
                '购买过多但销量持平导致库存不足',
            ('IGUAL', 'EXCESO', 'OPTIMO'):
                '购买过多但销量正好抵消',
            ('IGUAL', 'EXCESO', 'EXCEDENTE'):
                '购买过多且销量持平导致库存过剩',
            ('IGUAL', 'IGUAL', 'DEFICIT'):
                '销售和购买相同但库存不足',
            ('IGUAL', 'IGUAL', 'OPTIMO'):
                '销售和购买完美匹配，库存最佳',
            ('IGUAL', 'IGUAL', 'EXCEDENTE'):
                '销售和购买相同但库存过剩',
            ('IGUAL', 'DEFECTO', 'DEFICIT'):
                '购买不足导致库存不足',
            ('IGUAL', 'DEFECTO', 'OPTIMO'):
                '购买不足但销量正好保持库存',
            ('IGUAL', 'DEFECTO', 'EXCEDENTE'):
                '购买不足但仍有库存过剩',
            ('INFERIOR', 'EXCESO', 'DEFICIT'):
                '销量低且购买过多但库存仍不足',
            ('INFERIOR', 'EXCESO', 'OPTIMO'):
                '销量低但购买过多正好维持库存',
            ('INFERIOR', 'EXCESO', 'EXCEDENTE'):
                '销量低且购买过多导致库存过剩',
            ('INFERIOR', 'IGUAL', 'DEFICIT'):
                '销量低且购买未增加导致库存不足',
            ('INFERIOR', 'IGUAL', 'OPTIMO'):
                '销量低但购买正好维持库存',
            ('INFERIOR', 'IGUAL', 'EXCEDENTE'):
                '销量低且购买未增加导致库存过剩',
            ('INFERIOR', 'DEFECTO', 'DEFICIT'):
                '销量低且购买不足导致库存严重不足',
            ('INFERIOR', 'DEFECTO', 'OPTIMO'):
                '销量低且购买不足但库存仍最佳',
            ('INFERIOR', 'DEFECTO', 'EXCEDENTE'):
                '销量低且购买不足但仍有库存过剩',
        }
        
        clave = (
            escenario['ventas_vs_objetivo'],
            escenario['compras_vs_sugerido'],
            escenario['stock_vs_minimo']
        )
        
        escenario['descripcion'] = escenarios_descriptivos.get(
            clave, 
            f"Escenario: {escenario['codigo']}"
        )
        
        return escenario
    
    def aplicar_correccion_dataframe(
        self, 
        df: pd.DataFrame,
        columna_pedido: str = 'Unidades_Pedido',
        columna_stock_minimo: str = 'Stock_Minimo_Objetivo',
        columna_stock_real: str = 'Stock_Fisico',
        columna_categoria: str = 'Categoria',
        columna_ventas_reales: str = 'Unidades_Vendidas',
        columna_ventas_objetivo: str = 'Ventas_Objetivo',
        columna_compras_reales: str = 'Unidades_Recibidas',
        columna_compras_sugeridas: str = 'Unidades_Pedido'
    ) -> pd.DataFrame:
        """
        Aplica la corrección a todo un DataFrame de pedidos.
        
        Esta función aplica la fórmula de corrección a cada fila del DataFrame,
        actualizando el pedido teórico con el pedido corregido.
        
        Args:
            df (pd.DataFrame): DataFrame con los pedidos de FASE 1 y datos de corrección
            columna_pedido (str): Nombre de la columna con el pedido generado
            columna_stock_minimo (str): Nombre de la columna con el stock mínimo
            columna_stock_real (str): Nombre de la columna con el stock real
            columna_categoria (str): Nombre de la columna con la categoría ABC
            columna_ventas_reales (str): Nombre de la columna con ventas reales
            columna_ventas_objetivo (str): Nombre de la columna con ventas objetivo
            columna_compras_reales (str): Nombre de la columna con compras reales
            columna_compras_sugeridas (str): Nombre de la columna con compras sugeridas
        
        Returns:
            pd.DataFrame: DataFrame con columnas adicionales de corrección
        """
        logger.info("Aplicando corrección a DataFrame de pedidos...")
        
        df = df.copy()
        
        # Asegurar que existen las columnas necesarias
        cols_requeridas = [columna_pedido, columna_stock_real]
        for col in cols_requeridas:
            if col not in df.columns:
                df[col] = 0
        
        # Calcular stock mínimo si no existe
        if columna_stock_minimo not in df.columns:
            logger.debug("Calculando stock mínimo por categoría ABC...")
            df[columna_stock_minimo] = df.apply(
                lambda row: self.obtener_stock_minimo(
                    row.get(columna_categoria, 'C'),
                    row.get(columna_ventas_objetivo, 0),
                    row.get(columna_pedido, 0)
                ),
                axis=1
            )
        
        # Rellenar NaN en columnas numéricas
        df[columna_stock_minimo] = df[columna_stock_minimo].fillna(0)
        df[columna_stock_real] = df[columna_stock_real].fillna(0)
        
        # Calcular diferencia de stock
        df['Diferencia_Stock'] = df.apply(
            lambda row: self.calcular_diferencia_stock(
                row[columna_stock_minimo],
                row[columna_stock_real]
            ),
            axis=1
        )
        
        # Aplicar fórmula de corrección
        df['Pedido_Corregido'] = df.apply(
            lambda row: self.aplicar_formula_correccion(
                row[columna_pedido],
                row[columna_stock_minimo],
                row[columna_stock_real]
            ),
            axis=1
        )
        
        # Detectar escenario para cada artículo
        df['Escenario'] = df.apply(
            lambda row: self.detectar_escenario(
                pedido_generado=row[columna_pedido],
                stock_minimo=row[columna_stock_minimo],
                stock_real=row[columna_stock_real],
                ventas_reales=row.get(columna_ventas_reales, 0),
                ventas_objetivo=row.get(columna_ventas_objetivo, 0),
                compras_reales=row.get(columna_compras_reales, 0),
                compras_sugeridas=row.get(columna_compras_sugeridas, row[columna_pedido])
            )['codigo'],
            axis=1
        )
        
        # Añadir columna de razón de corrección
        df['Razon_Correccion'] = df.apply(
            lambda row: self._generar_razon_correccion(
                row[columna_stock_minimo],
                row[columna_stock_real],
                row['Diferencia_Stock'],
                row['Pedido_Corregido'],
                row[columna_pedido]
            ),
            axis=1
        )
        
        # Calcular métricas de corrección
        articulos_corregidos = len(df[df['Pedido_Corregido'] != df[columna_pedido]])
        articulos_aumentados = len(df[df['Pedido_Corregido'] > df[columna_pedido]])
        articulos_reducidos = len(df[df['Pedido_Corregido'] < df[columna_pedido]])
        
        logger.info(f"Corrección completada:")
        logger.info(f"  - Artículos corregidos: {articulos_corregidos}/{len(df)}")
        logger.info(f"  - Artículos aumentados: {articulos_aumentados}")
        logger.info(f"  - Artículos reducidos: {articulos_reducidos}")
        
        return df
    
    def _generar_razon_correccion(
        self,
        stock_minimo: float,
        stock_real: float,
        diferencia_stock: float,
        pedido_corregido: float,
        pedido_original: float
    ) -> str:
        """
        Genera una explicación legible de la corrección aplicada.
        
        Args:
            stock_minimo (float): Stock mínimo objetivo
            stock_real (float): Stock real actual
            diferencia_stock (float): Diferencia calculada
            pedido_corregido (float): Pedido resultante
            pedido_original (float): Pedido original de FASE 1
        
        Returns:
            str: Descripción de la corrección aplicada
        """
        if pedido_corregido == pedido_original:
            return "Sin corrección necesaria"
        
        if stock_real >= stock_minimo:
            if stock_real > stock_minimo:
                exceso = stock_real - stock_minimo
                return f"Reducir {exceso:.0f} unidades (stock excedente)"
            else:
                return "Mantener pedido (stock óptimo)"
        else:
            deficit = stock_minimo - stock_real
            return f"Aumentar {deficit:.0f} unidades (recuperar stock mínimo)"
    
    def calcular_metricas_correccion(
        self, 
        df: pd.DataFrame,
        columna_pedido_original: str = 'Unidades_Pedido',
        columna_pedido_corregido: str = 'Pedido_Corregido',
        columna_ventas_reales: str = 'Unidades_Vendidas',
        columna_ventas_objetivo: str = 'Ventas_Objetivo'
    ) -> Dict[str, Any]:
        """
        Calcula métricas de evaluación del sistema de corrección.
        
        Args:
            df (pd.DataFrame): DataFrame con pedidos corregidos
            columna_pedido_original (str): Nombre de la columna con pedido original
            columna_pedido_corregido (str): Nombre de la columna con pedido corregido
            columna_ventas_reales (str): Nombre de la columna con ventas reales
            columna_ventas_objetivo (str): Nombre de la columna con ventas objetivo
        
        Returns:
            Dict[str, Any]: Métricas calculadas
        """
        metricas = {}
        
        # Total de artículos
        metricas['total_articulos'] = len(df)
        
        # Artículos corregidos
        df_correccion = df[df[columna_pedido_corregido] != df[columna_pedido_original]]
        metricas['articulos_corregidos'] = len(df_correccion)
        metricas['porcentaje_corregidos'] = (
            metricas['articulos_corregidos'] / metricas['total_articulos'] * 100
            if metricas['total_articulos'] > 0 else 0
        )
        
        # Cambios de cantidad
        cambios = df_correccion[columna_pedido_corregido] - df_correccion[columna_pedido_original]
        metricas['total_unidades_original'] = df[columna_pedido_original].sum()
        metricas['total_unidades_corregido'] = df[columna_pedido_corregido].sum()
        metricas['diferencia_unidades'] = (
            metricas['total_unidades_corregido'] - metricas['total_unidades_original']
        )
        metricas['porcentaje_cambio'] = (
            metricas['diferencia_unidades'] / metricas['total_unidades_original'] * 100
            if metricas['total_unidades_original'] > 0 else 0
        )
        
        # Índice de precisión (si hay ventas reales)
        if columna_ventas_reales in df.columns and df[columna_ventas_reales].sum() > 0:
            ventas_totales = df[columna_ventas_reales].sum()
            ventas_objetivo = df[columna_ventas_objetivo].sum()
            metricas['precision_forecast'] = ventas_totales / ventas_objetivo if ventas_objetivo > 0 else 1.0
        else:
            metricas['precision_forecast'] = None
        
        # Distribución por escenario
        metricas['distribucion_escenarios'] = df['Escenario'].value_counts().to_dict()
        
        # Artículos en alerta de stock bajo
        umbral = self.config.umbral_alerta_stock
        if umbral > 0:
            metricas['articulos_alerta_stock'] = len(
                df[df['Stock_Fisico'] <= umbral]
            )
        else:
            metricas['articulos_alerta_stock'] = 0
        
        return metricas
    
    def generar_alertas(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """
        Genera alertas para situaciones que requieren atención.
        
        Args:
            df (pd.DataFrame): DataFrame con datos de corrección
        
        Returns:
            List[Dict[str, Any]]: Lista de alertas generadas
        """
        alertas = []
        
        # Artículos con stock crítico (menor o igual a 0)
        stock_cero = df[df['Stock_Fisico'] <= 0]
        if len(stock_cero) > 0:
            alertas.append({
                'tipo': 'STOCK_CRITICO',
                'nivel': 'ALTO',
                'mensaje': f'{len(stock_cero)} artículos con stock en 0 o negativo',
                'articulos': stock_cero['Codigo_Articulo'].head(10).tolist() if 'Codigo_Articulo' in stock_cero.columns else []
            })
        
        # Artículos con corrección significativa (más del 50%)
        df['pct_cambio'] = abs(df['Pedido_Corregido'] - df['Unidades_Pedido']) / df['Unidades_Pedido'].replace(0, 1)
        cambios_significativos = df[df['pct_cambio'] > 0.5]
        if len(cambios_significativos) > 0:
            alertas.append({
                'tipo': 'CAMBIOS_SIGNIFICATIVOS',
                'nivel': 'MEDIO',
                'mensaje': f'{len(cambios_significativos)} artículos con cambios superiores al 50%',
                'articulos': cambios_significativos['Codigo_Articulo'].head(10).tolist() if 'Codigo_Articulo' in cambios_significativos.columns else []
            })
        
        # Artículos sin ventas pero con stock
        sin_ventas_con_stock = df[(df['Unidades_Vendidas'] == 0) & (df['Stock_Fisico'] > 0)]
        if len(sin_ventas_con_stock) > 0:
            alertas.append({
                'tipo': 'SIN_VENTAS',
                'nivel': 'BAJO',
                'mensaje': f'{len(sin_ventas_con_stock)} artículos con stock pero sin ventas',
                'articulos': sin_ventas_con_stock['Codigo_Articulo'].head(10).tolist() if 'Codigo_Articulo' in sin_ventas_con_stock.columns else []
            })
        
        return alertas


# Funciones de utilidad
def crear_correction_engine(
    config_abc: Optional[Dict[str, Any]] = None,
    politica_stock_minimo: Optional[Dict[str, float]] = None
) -> CorrectionEngine:
    """
    Crea una instancia del CorrectionEngine con configuración estándar.
    
    Args:
        config_abc (Optional[Dict]): Configuración ABC del sistema
        politica_stock_minimo (Optional[Dict]): Política de stock mínimo por categoría
    
    Returns:
        CorrectionEngine: Instancia del motor de corrección
    """
    config = ConfiguracionCorreccion(
        politica_stock_minimo=politica_stock_minimo or {
            'A': 1.5,
            'B': 1.0,
            'C': 0.5,
            'D': 0.0
        },
        umbral_alerta_stock=5,
        permitir_pedidos_negativos=False
    )
    
    return CorrectionEngine(config_abc=config_abc, configuracion=config)


if __name__ == "__main__":
    # Ejemplo de uso y prueba de la fórmula
    print("CorrectionEngine - Motor de corrección FASE 2")
    print("=" * 60)
    print("\nVerificación de la fórmula contra escenarios del documento OPCIONES:")
    print("-" * 60)
    
    # Crear motor
    engine = crear_correction_engine()
    
    # Casos de prueba basados en el documento OPCIONES
    casos_prueba = [
        # (Stock_Inicial, Stock_Min, Ventas_Obj, Pedido_Orig, Compra_Real, Venta_Real, Stock_Teo, Stock_Real, Esperado)
        (30, 20, 20, 10, 30, 50, 10, 30, 0),    # Excedente de compras, ventas altas, stock real mayor que mínimo
        (30, 20, 20, 10, 30, 50, 10, 20, 10),   # Excedente de compras, ventas altas, stock real = mínimo
        (30, 20, 20, 10, 30, 50, 10, 10, 20),   # Excedente de compras, ventas altas, stock real menor que mínimo
        (30, 20, 20, 10, 10, 30, 10, 10, 20),   # Compras justas, ventas altas, stock real menor que mínimo
        (30, 20, 20, 10, 10, 20, 20, 20, 10),   # Compras justas, ventas exactas, stock real = mínimo
        (30, 20, 20, 10, 10, 10, 30, 10, 10),   # Compras justas, ventas bajas, stock real menor que mínimo
        (30, 20, 20, 10, 5, 30, 5, 10, 20),     # Compras bajas, ventas altas, stock real menor que mínimo
        (30, 20, 20, 10, 5, 15, 20, 20, 10),     # Compras bajas, ventas bajas, stock real = mínimo
    ]
    
    print(f"{'Pedido':>8} | {'Stock_Min':>10} | {'Stock_Real':>10} | {'Esperado':>10} | {'Calculado':>10} | {'Match':>6}")
    print("-" * 70)
    
    todos_correctos = True
    for i, caso in enumerate(casos_prueba):
        _, stock_min, _, pedido_orig, _, _, _, stock_real, esperado = caso
        
        resultado = engine.aplicar_formula_correccion(pedido_orig, stock_min, stock_real)
        match = abs(resultado - esperado) < 0.01
        todos_correctos = todos_correctos and match
        
        print(f"{pedido_orig:>8} | {stock_min:>10} | {stock_real:>10} | {esperado:>10.0f} | {resultado:>10.0f} | {'OK':>6}" if match else f"{pedido_orig:>8} | {stock_min:>10} | {stock_real:>10} | {esperado:>10.0f} | {resultado:>10.0f} | {'FAIL':>6}")
    
    print("-" * 70)
    if todos_correctos:
        print("✓ Todos los escenarios verificados correctamente")
    else:
        print("✗ Algunos escenarios no coinciden")
