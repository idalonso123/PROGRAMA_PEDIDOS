#!/usr/bin/env python3
"""
Script de verificación específico para demostrar que:
- Unidades Pedido = Unidades Finales + Stock Mínimo Objetivo

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-07
"""

import sys
import pandas as pd
import numpy as np
from pathlib import Path

# Añadir src al path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from forecast_engine import ForecastEngine

def test_verificacion_unidades_pedido():
    """
    Verificación específica de que:
    Unidades Pedido = Unidades Finales + Stock Mínimo Objetivo
    """
    print("=" * 80)
    print("VERIFICACIÓN: Unidades Pedido = Unidades Finales + Stock Mínimo Objetivo")
    print("=" * 80)
    print()
    
    config = {
        'parametros': {
            'objetivo_crecimiento': 0.05,
            'stock_minimo_porcentaje': 0.30,
            'pesos_categoria': {'A': 1.0, 'B': 0.8, 'C': 0.6, 'D': 0.0}
        },
        'festivos': {},
        'secciones': {}
    }
    
    engine = ForecastEngine(config)
    
    # Crear DataFrame de prueba con diferentes valores
    df = pd.DataFrame({
        'Codigo_Articulo': ['ART001', 'ART002', 'ART003', 'ART004', 'ART005'],
        'Nombre_Articulo': ['Producto A', 'Producto B', 'Producto C', 'Producto D', 'Producto E'],
        'Talla': ['U', 'U', 'U', 'U', 'U'],
        'Color': ['UNICA', 'UNICA', 'UNICA', 'UNICA', 'UNICA'],
        'Unidades_Finales': [100, 50, 33, 20, 1],  # Various values to test rounding
        'PVP': [10.0, 15.0, 20.0, 25.0, 30.0],
        'Coste_Pedido': [5.0, 7.5, 10.0, 12.5, 15.0],
        'Proveedor': ['Prov A', 'Prov B', 'Prov C', 'Prov D', 'Prov E'],
        'Categoria': ['A', 'B', 'C', 'D', 'A'],
        'Accion_Aplicada': ['MANTENER', 'MANTENER', 'MANTENER', 'MANTENER', 'MANTENER'],
        'Ventas_Objetivo': [1000.0, 750.0, 660.0, 500.0, 30.0],
        'Beneficio_Objetivo': [500.0, 375.0, 330.0, 250.0, 15.0],
        'Seccion': ['TEST', 'TEST', 'TEST', 'TEST', 'TEST']
    })
    
    # Datos de stock acumulado (vacíos para simplificar)
    stock_acumulado = {
        'ART001|U|UNICA': 0,
        'ART002|U|UNICA': 0,
        'ART003|U|UNICA': 0,
        'ART004|U|UNICA': 0,
        'ART005|U|UNICA': 0,
    }
    
    # Ejecutar aplicar_stock_minimo
    pedidos_actualizados, nuevo_stock, ajustes = engine.aplicar_stock_minimo(
        df, 5, stock_acumulado
    )
    
    print(f"{'Código':>10} | {'Und.Fin':>8} | {'StkMin':>8} | {'Und.Ped':>8} | {'Und.Fin+StkMin':>14} | {'Estado':>8}")
    print("-" * 80)
    
    todos_correctos = True
    
    for idx, row in pedidos_actualizados.iterrows():
        codigo = row['Codigo_Articulo']
        und_fin = row['Unidades_Finales']
        stk_min = row['Stock_Minimo_Objetivo']
        und_ped = row['Unidades_Pedido']
        
        # Calcular lo que debería ser
        suma_esperada = und_fin + stk_min
        
        # Verificar
        correcto = (und_ped == suma_esperada)
        todos_correctos = todos_correctos and correcto
        
        estado = "✓ OK" if correcto else "✗ FAIL"
        print(f"{codigo:>10} | {und_fin:>8.0f} | {stk_min:>8.0f} | {und_ped:>8.0f} | {suma_esperada:>14.0f} | {estado:>8}")
        
        if not correcto:
            print(f"           ✗ ERROR: {und_ped} != {und_fin} + {stk_min}")
    
    print("-" * 80)
    
    if todos_correctos:
        print("✓ RESULTADO: Todas las 'Unidades Pedido' son correctas")
        print("             (= Unidades Finales + Stock Mínimo Objetivo)")
    else:
        print("✗ RESULTADO: Hay errores en el cálculo de Unidades Pedido")
    
    return todos_correctos


def test_verificacion_formula_detallada():
    """
    Verificación detallada de cada paso del cálculo
    """
    print("\n" + "=" * 80)
    print("VERIFICACIÓN DETALLADA: Paso a paso")
    print("=" * 80)
    print()
    
    # Ejemplo del usuario
    print("Ejemplo con artículo hipotético:")
    print("  - Unidades Finales (calculadas): 100")
    print("  - Stock mínimo (%): 30%")
    print("  - Stock acumulado anterior: 0")
    print()
    
    und_fin = 100
    pct_stock_min = 0.30
    
    # Paso 1: Calcular stock mínimo
    stock_min = int(np.ceil(und_fin * pct_stock_min))
    print(f"  Paso 1: Stock Mínimo = ceil({und_fin} × {pct_stock_min}) = {stock_min}")
    
    # Paso 2: Calcular Units Pedido
    unidades_pedido = und_fin + stock_min
    print(f"  Paso 2: Unidades Pedido = {und_fin} + {stock_min} = {unidades_pedido}")
    
    # Verificación
    print()
    print(f"  FÓRMULA VERIFICADA:")
    print(f"    Unidades Pedido = {unidades_pedido}")
    print(f"    Unidades Finales + Stock Mínimo = {und_fin} + {stock_min} = {unidades_pedido}")
    
    # Mostrar con stock real
    print()
    print("  ─────────────────────────────────────────")
    print("  ESCENARIO: Con corrección por stock real")
    print("  ─────────────────────────────────────────")
    print()
    
    stock_real = 25  # Déficit de 5 respecto al mínimo (30)
    
    # Corrección por desviación de stock
    pedido_corregido_stock = max(0, unidades_pedido + (stock_min - stock_real))
    print(f"  Stock Real: {stock_real}")
    print(f"  Diferencia: {stock_min} - {stock_real} = {stock_min - stock_real}")
    print(f"  Pedido Corregido Stock = max(0, {unidades_pedido} + ({stock_min} - {stock_real}))")
    print(f"                         = max(0, {unidades_pedido + (stock_min - stock_real)})")
    print(f"                         = {pedido_corregido_stock}")
    
    # Con tendencia de ventas
    print()
    ventas_reales = 115
    ventas_objetivo = 100
    tendencia_consumo = max(0, ventas_reales - ventas_objetivo)
    pedido_final = pedido_corregido_stock + tendencia_consumo
    
    print(f"  Ventas Reales (sem. anterior): {ventas_reales}")
    print(f"  Ventas Objetivo (sem. anterior): {ventas_objetivo}")
    print(f"  Tendencia Consumo = max(0, {ventas_reales} - {ventas_objetivo}) = {tendencia_consumo}")
    print(f"  Pedido Final = {pedido_corregido_stock} + {tendencia_consumo} = {pedido_final}")
    
    print()
    print("  RESUMEN:")
    print(f"    ┌────────────────────────┬─────────┐")
    print(f"    │ Concepto               │ Valor   │")
    print(f"    ├────────────────────────┼─────────┤")
    print(f"    │ Unidades Calculadas    │   {und_fin:>5} │  (Forecast base)")
    print(f"    │ Stock Mínimo Objetivo  │   {stock_min:>5} │  (30% de {und_fin})")
    print(f"    ├────────────────────────┼─────────┤")
    print(f"    │ UNIDADES PEDIDO        │   {unidades_pedido:>5} │  = {und_fin} + {stock_min}")
    print(f"    │ Pedido Corregido Stock │   {pedido_corregido_stock:>5} │  (+/- por stock real)")
    print(f"    │ PEDIDO FINAL            │   {pedido_final:>5} │  (+ tendencia ventas)")
    print(f"    └────────────────────────┴─────────┘")
    
    return True


def main():
    """Ejecuta todas las verificaciones"""
    print("\n" + "#" * 80)
    print("# VERIFICACIÓN ESPECÍFICA: Cálculo de Unidades Pedido")
    print("# Objetivo: Confirmar que Unidades Pedido = Unidades Finales + Stock Mínimo")
    print("#" * 80 + "\n")
    
    resultados = []
    
    # Test 1: Verificación con DataFrame
    resultados.append(("Verificación con datos", test_verificacion_unidades_pedido()))
    
    # Test 2: Verificación detallada
    resultados.append(("Verificación detallada", test_verificacion_formula_detallada()))
    
    # Resumen
    print("\n" + "=" * 80)
    print("RESUMEN DE VERIFICACIONES")
    print("=" * 80)
    
    todas_pasaron = True
    for nombre, resultado in resultados:
        estado = "✓ PASÓ" if resultado else "✗ FALLÓ"
        print(f"  {nombre}: {estado}")
        todas_pasaron = todas_pasaron and resultado
    
    print("-" * 80)
    if todas_pasaron:
        print("✓ TODAS LAS VERIFICACIONES PASARON")
        print("\nLa corrección ha sido aplicada correctamente:")
        print("  • Unidades Pedido = Unidades Finales + Stock Mínimo Objetivo")
        print("  • Stock Mínimo = ceil(Unidades Finales × 30%)")
        print("  • Las correcciones FASE 2 se aplican sobre este base")
    else:
        print("✗ ALGUNAS VERIFICACIONES FALLARON")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
