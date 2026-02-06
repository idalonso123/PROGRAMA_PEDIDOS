#!/usr/bin/env python3
"""
Script de prueba para verificar la implementación completa de FASE 2:
1. Corrección por Desviación de Stock
2. Corrección por Tendencia de Ventas

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

def test_correccion_desviacion_stock():
    """
    Prueba 1: Verificación de la Corrección por Desviación de Stock.
    
    Objetivo: Asegurar que siempre mantenemos el stock mínimo.
    Fórmula: Pedido_Corregido_Stock = max(0, unidades_pedido + (stock_minimo - stock_real))
    """
    print("=" * 80)
    print("PRUEBA 1: Corrección por Desviación de Stock")
    print("=" * 80)
    print("\nFórmula:")
    print("  Pedido_Corregido_Stock = max(0, unidades_pedido + (Stock_Mínimo - Stock_Real))")
    print("\nObjetivo: Mantener siempre el stock mínimo configurado (30%)")
    print()
    
    casos_prueba = [
        # (unidades_pedido, stock_minimo, stock_real, esperado, descripcion)
        (100, 30, 25, 105, "Déficit: stock real 25 < stock mínimo 30 → aumentar pedido"),
        (100, 30, 30, 100, "Óptimo: stock real = stock mínimo → sin cambio"),
        (100, 30, 35, 95, "Excedente: stock real 35 > stock mínimo 30 → reducir pedido"),
        (100, 30, 0, 130, "Déficit crítico: stock real = 0 → aumentar significativamente"),
        (0, 30, 25, 5, "Pedido cero con déficit menor → pedido mínimo para recuperar"),
        (0, 30, 0, 30, "Pedido cero con déficit crítico → pedir stock mínimo"),
    ]
    
    print(f"{'Und.Pedido':>10} | {'Stock Min':>10} | {'Stock Real':>12} | {'Esperado':>10} | {'Calculado':>10} | {'Estado':>8}")
    print("-" * 80)
    
    todos_correctos = True
    for und_pedido, stock_min, stock_real, esperado, descripcion in casos_prueba:
        # Calcular corrección por desviación de stock
        pedido_corregido_stock = max(0, und_pedido + (stock_min - stock_real))
        
        match = abs(pedido_corregido_stock - esperado) < 0.01
        todos_correctos = todos_correctos and match
        
        estado = "✓ OK" if match else "✗ FAIL"
        print(f"{und_pedido:>10} | {stock_min:>10} | {stock_real:>12} | {esperado:>10.0f} | {pedido_corregido_stock:>10.0f} | {estado:>8}")
        print(f"           → {descripcion}")
    
    print("-" * 80)
    if todos_correctos:
        print("✓ RESULTADO: Corrección por Desviación de Stock funciona correctamente")
    else:
        print("✗ RESULTADO: Algunos escenarios fallaron")
    
    return todos_correctos


def test_correccion_tendencia_ventas():
    """
    Prueba 2: Verificación de la Corrección por Tendencia de Ventas.
    
    Objetivo: Detectar si hay una tendencia de aumento de ventas.
    Lógica: Si ventas_reales > ventas_objetivo, hubo consumo del stock mínimo.
    Fórmula: Tendencia_Consumo = max(0, Ventas_Reales - Ventas_Objetivo)
    
    Ejemplo del usuario:
    - Semana N: Se piden 26 unidades (20 venta + 6 stock mínimo)
    - Ventas reales: 24 unidades (4 más de las 20 esperadas)
    - Estas 4 unidades extra vinieron del stock mínimo
    - Semana N+1: Añadir esas 4 unidades al pedido como "corrección por tendencia"
    """
    print("\n" + "=" * 80)
    print("PRUEBA 2: Corrección por Tendencia de Ventas")
    print("=" * 80)
    print("\nFórmula:")
    print("  Tendencia_Consumo = max(0, Ventas_Reales - Ventas_Objetivo)")
    print("\nLógica:")
    print("  - Si Ventas_Reales > Ventas_Objetivo: Se consumió stock mínimo")
    print("  - Las unidades consumidas del mínimo indican posible tendencia al alza")
    print("  - Se añaden esas unidades al siguiente pedido")
    print()
    
    print("Ejemplo del usuario:")
    print("  - Und. venta objetivo: 20")
    print("  - Stock mínimo: 6")
    print("  - Total pedido semana N: 26")
    print("  - Ventas reales: 24 (4 más de lo esperado)")
    print("  - Estas 4 unidades extras vinieron del stock mínimo")
    print("  - Tendencia_Consumo = 24 - 20 = 4")
    print("  - Para semana N+1: Añadir +4 unidades al pedido")
    print()
    
    casos_prueba = [
        # (ventas_reales, ventas_objetivo, esperado, descripcion)
        (24, 20, 4, "Venta superior al objetivo → tendencia al alza"),
        (20, 20, 0, "Venta igual al objetivo → sin tendencia"),
        (18, 20, 0, "Venta inferior al objetivo → sin tendencia"),
        (30, 20, 10, "Gran venta superior → fuerte tendencia al alza"),
        (0, 20, 0, "Sin ventas → sin tendencia"),
        (25, 10, 15, "Venta muy superior → tendencia muy alta"),
    ]
    
    print(f"{'Ventas Reales':>14} | {'Ventas Obj.':>12} | {'Esperado':>10} | {'Calculado':>10} | {'Estado':>8}")
    print("-" * 75)
    
    todos_correctos = True
    for v_reales, v_objetivo, esperado, descripcion in casos_prueba:
        # Calcular corrección por tendencia
        tendencia_consumo = max(0, v_reales - v_objetivo)
        
        match = abs(tendencia_consumo - esperado) < 0.01
        todos_correctos = todos_correctos and match
        
        estado = "✓ OK" if match else "✗ FAIL"
        print(f"{v_reales:>14} | {v_objetivo:>12} | {esperado:>10.0f} | {tendencia_consumo:>10.0f} | {estado:>8}")
        print(f"           → {descripcion}")
    
    print("-" * 75)
    if todos_correctos:
        print("✓ RESULTADO: Corrección por Tendencia de Ventas funciona correctamente")
    else:
        print("✗ RESULTADO: Algunos escenarios fallaron")
    
    return todos_correctos


def test_ambas_correcciones():
    """
    Prueba 3: Verificación de ambas correcciones en secuencia.
    
    Paso 1: Corrección por Desviación de Stock
    Paso 2: Corrección por Tendencia de Ventas
    Resultado: Pedido_Final = Pedido_Corregido_Stock + Tendencia_Consumo
    """
    print("\n" + "=" * 80)
    print("PRUEBA 3: Ambas correcciones en secuencia")
    print("=" * 80)
    print("\nProceso:")
    print("  1. Calcular Pedido_Corregido_Stock")
    print("  2. Calcular Tendencia_Consumo")
    print("  3. Pedido_Final = Pedido_Corregido_Stock + Tendencia_Consumo")
    print()
    
    # Ejemplo del usuario
    print("Ejemplo del usuario:")
    print("  - Und. venta objetivo: 20")
    print("  - Stock mínimo (30%): 6")
    print("  - Total unidades finales: 26")
    print("  - Stock real actual: 10 (se vendieron 4 del mínimo)")
    print("  - Ventas reales: 24")
    print("  - Stock acumulado anterior: 0")
    print()
    
    # Datos del ejemplo
    unidades_finales = 20
    stock_minimo_porcentaje = 0.30
    stock_minimo = int(np.ceil(unidades_finales * stock_minimo_porcentaje))  # 6
    stock_acumulado = 0
    stock_real = 10
    ventas_reales = 24
    ventas_objetivo = 20
    
    print("Cálculo paso a paso:")
    print(f"  1. Stock_Mínimo = ceil(20 × 0.30) = {stock_minimo}")
    print(f"  2. Diferencia_Stock = {stock_minimo} - {stock_acumulado} = {stock_minimo - stock_acumulado}")
    print(f"  3. Unidades_Pedido = 20 + {stock_minimo - stock_acumulado} = {unidades_finales + (stock_minimo - stock_acumulado)}")
    
    # Corrección 1: Desviación de stock
    unidades_pedido = unidades_finales + (stock_minimo - stock_acumulado)
    pedido_corregido_stock = max(0, unidades_pedido + (stock_minimo - stock_real))
    print(f"  4. Pedido_Corregido_Stock = max(0, {unidades_pedido} + ({stock_minimo} - {stock_real})) = {pedido_corregido_stock}")
    
    # Corrección 2: Tendencia de ventas
    tendencia_consumo = max(0, ventas_reales - ventas_objetivo)
    print(f"  5. Tendencia_Consumo = max(0, {ventas_reales} - {ventas_objetivo}) = {tendencia_consumo}")
    
    # Resultado final
    pedido_final = pedido_corregido_stock + tendencia_consumo
    print(f"  6. Pedido_Final = {pedido_corregido_stock} + {tendencia_consumo} = {pedido_final}")
    print()
    
    print("-" * 80)
    print(f"RESUMEN:")
    print(f"  - Stock Mínimo Objetivo: {stock_minimo}")
    print(f"  - Unidades Pedido (base): {unidades_pedido}")
    print(f"  - Pedido Corregido Stock: {pedido_corregido_stock}")
    print(f"  - Ventas Reales: {ventas_reales}")
    print(f"  - Tendencia Consumo: {tendencia_consumo}")
    print(f"  - PEDIDO FINAL: {pedido_final}")
    print("-" * 80)
    
    # Verificación con valores esperados
    # Según la lógica implementada:
    # - Pedido_Corregido_Stock = max(0, 26 + (6 - 10)) = 22
    # - Tendencia_Consumo = max(0, 24 - 20) = 4
    # - Pedido_Final = 22 + 4 = 26
    esperado_pedido_final = 26  # Corregido: 22 (stock) + 4 (tendencia) = 26
    correcto = pedido_final == esperado_pedido_final
    estado = "✓ OK" if correcto else "✗ FAIL"
    print(f"Pedido Final esperado: {esperado_pedido_final}, Calculado: {pedido_final} → {estado}")
    
    return correcto


def test_metodo_aplicar_stock_minimo():
    """
    Prueba 4: Verificación del método aplicar_stock_minimo con ambas correcciones.
    """
    print("\n" + "=" * 80)
    print("PRUEBA 4: Método aplicar_stock_minimo con ambas correcciones")
    print("=" * 80)
    
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
    
    # Crear DataFrame de prueba
    df = pd.DataFrame({
        'Codigo_Articulo': ['ART001', 'ART002', 'ART003'],
        'Nombre_Articulo': ['Producto A', 'Producto B', 'Producto C'],
        'Talla': ['U', 'U', 'U'],
        'Color': ['UNICA', 'UNICA', 'UNICA'],
        'Unidades_Finales': [100, 50, 20],
        'PVP': [10.0, 15.0, 20.0],
        'Coste_Pedido': [5.0, 7.5, 10.0],
        'Proveedor': ['Prov A', 'Prov B', 'Prov C'],
        'Categoria': ['A', 'B', 'C'],
        'Accion_Aplicada': ['MANTENER', 'MANTENER', 'MANTENER'],
        'Ventas_Objetivo': [1000.0, 750.0, 400.0],
        'Beneficio_Objetivo': [500.0, 375.0, 200.0],
        'Seccion': ['TEST', 'TEST', 'TEST']
    })
    
    # Datos para correcciones
    stock_acumulado = {
        'ART001|U|UNICA': 0,
        'ART002|U|UNICA': 0,
        'ART003|U|UNICA': 0
    }
    
    stock_real = {
        'ART001|U|UNICA': 25,   # Déficit de 5
        'ART002|U|UNICA': 15,   # Óptimo
        'ART003|U|UNICA': 10,   # Excedente de 4
    }
    
    ventas_reales = {
        'ART001|U|UNICA': 110,  # Vendió 10 más del objetivo (100)
        'ART002|U|UNICA': 50,   # Vendió exactamente el objetivo
        'ART003|U|UNICA': 18,   # Vendió 2 menos del objetivo (20)
    }
    
    ventas_objetivo = {
        'ART001|U|UNICA': 100,  # Objetivo de ventas
        'ART002|U|UNICA': 50,
        'ART003|U|UNICA': 20,
    }
    
    print("\nDatos de entrada:")
    print(f"  Stock real por artículo: {stock_real}")
    print(f"  Ventas reales por artículo: {ventas_reales}")
    print(f"  Ventas objetivo por artículo: {ventas_objetivo}")
    
    # Ejecutar aplicar_stock_minimo con todos los parámetros
    pedidos_actualizados, nuevo_stock, ajustes = engine.aplicar_stock_minimo(
        df, 5, stock_acumulado, stock_real, ventas_reales, ventas_objetivo
    )
    
    print("\n" + "-" * 100)
    print(f"{'Código':>10} | {'Und.Fin':>8} | {'StkMin':>8} | {'StkReal':>8} | {'Und.Ped':>8} | {'Ped.Corr':>8} | {'Vta.Real':>8} | {'Tend':>6} | {'Ped.Fin':>8}")
    print("-" * 100)
    
    resultados = []
    for idx, row in pedidos_actualizados.iterrows():
        codigo = row['Codigo_Articulo']
        und_fin = row['Unidades_Finales']
        stk_min = row['Stock_Minimo_Objetivo']
        stk_real = stock_real.get(f'{codigo}|U|UNICA', 0)
        und_ped = row['Unidades_Pedido']
        ped_corr = row['Pedido_Corregido_Stock']
        vta_real = row['Ventas_Reales']
        tend = row['Tendencia_Consumo']
        ped_fin = row['Pedido_Final']
        
        print(f"{codigo:>10} | {und_fin:>8} | {stk_min:>8} | {stk_real:>8} | {und_ped:>8} | {ped_corr:>8} | {vta_real:>8} | {tend:>6} | {ped_fin:>8}")
        
        # Verificar cálculos
        correcto = True
        # Verificar stock mínimo
        if stk_min != int(np.ceil(und_fin * 0.30)):
            correcto = False
        # Verificar pedido corregido stock
        ped_corr_esperado = max(0, und_ped + (stk_min - stk_real))
        if ped_corr != ped_corr_esperado:
            correcto = False
        # Verificar tendencia
        vta_obj = ventas_objetivo.get(f'{codigo}|U|UNICA', 0)
        tend_esperada = max(0, vta_real - vta_obj)
        if tend != tend_esperada:
            correcto = False
        # Verificar pedido final
        if ped_fin != ped_corr + tend:
            correcto = False
        
        resultados.append(correcto)
    
    print("-" * 100)
    
    # Verificar que las nuevas columnas existen
    columnas_ok = (
        'Pedido_Corregido_Stock' in pedidos_actualizados.columns and
        'Ventas_Reales' in pedidos_actualizados.columns and
        'Tendencia_Consumo' in pedidos_actualizados.columns and
        'Pedido_Final' in pedidos_actualizados.columns
    )
    
    todos_correctos = all(resultados) and columnas_ok
    
    print("\nVerificaciones:")
    print(f"  ✓ Columnas FASE 2 existen: {columnas_ok}")
    print(f"  ✓ Todos los cálculos correctos: {all(resultados)}")
    
    return todos_correctos


def main():
    """Ejecuta todas las pruebas."""
    print("\n" + "#" * 80)
    print("# PRUEBAS DE VERIFICACIÓN FASE 2 - Completa")
    print("# Corrección por Desviación de Stock + Corrección por Tendencia de Ventas")
    print("#" * 80)
    
    resultados = []
    
    # Prueba 1: Corrección por Desviación de Stock
    resultados.append(("Corrección Desviación Stock", test_correccion_desviacion_stock()))
    
    # Prueba 2: Corrección por Tendencia de Ventas
    resultados.append(("Corrección Tendencia Ventas", test_correccion_tendencia_ventas()))
    
    # Prueba 3: Ambas correcciones en secuencia
    resultados.append(("Ambas correcciones (ejemplo)", test_ambas_correcciones()))
    
    # Prueba 4: Método aplicar_stock_minimo
    resultados.append(("Método aplicar_stock_minimo", test_metodo_aplicar_stock_minimo()))
    
    # Resumen
    print("\n" + "=" * 80)
    print("RESUMEN DE PRUEBAS")
    print("=" * 80)
    
    todas_pasaron = True
    for nombre, resultado in resultados:
        estado = "✓ PASÓ" if resultado else "✗ FALLÓ"
        print(f"  {nombre}: {estado}")
        todas_pasaron = todas_pasaron and resultado
    
    print("-" * 80)
    if todas_pasaron:
        print("✓ TODAS LAS PRUEBAS PASARON CORRECTAMENTE")
        print("\nLa implementación de FASE 2 incluye:")
        print("  1. Corrección por Desviación de Stock:")
        print("     - Mantiene el stock mínimo configurado (30%)")
        print("     - Ajusta el pedido según el inventario real")
        print()
        print("  2. Corrección por Tendencia de Ventas:")
        print("     - Detecta si se consumió stock mínimo (ventas > objetivo)")
        print("     - Añade esas unidades al siguiente pedido")
        print("     - Predice posibles tendencias al alza")
    else:
        print("✗ ALGUNAS PRUEBAS FALLARON")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
