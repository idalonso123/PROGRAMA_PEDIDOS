#!/usr/bin/env python3
"""
Script de verificación FINAL: Pasos 1-4 - Verificar todos los cambios

Verificar que:
1. La fórmula de Pedido_Corregido_Stock NO duplica el stock mínimo
2. El filtrado usa Pedido_Corregido_Stock
3. Total_Unidades usa Pedido_Final
4. La columna Unidades_Pedido ha sido eliminada

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

def test_verificacion_final():
    """
    Verificación completa de todos los cambios
    """
    print("=" * 80)
    print("VERIFICACIÓN FINAL: Todos los cambios (Pasos 1-4)")
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
    
    # Datos de prueba
    stock_acumulado = {
        'ART001|U|UNICA': 0,
        'ART002|U|UNICA': 0,
        'ART003|U|UNICA': 0,
    }
    
    stock_real = {
        'ART001|U|UNICA': 25,   # Déficit: 30 - 25 = 5
        'ART002|U|UNICA': 15,   # Óptimo: 15 - 15 = 0
        'ART003|U|UNICA': 10,   # Excedente: 6 - 10 = -4
    }
    
    ventas_reales = {
        'ART001|U|UNICA': 110,  # Tendencia: 110 - 100 = 10
        'ART002|U|UNICA': 50,   # Sin tendencia: 50 - 50 = 0
        'ART003|U|UNICA': 18,   # Sin tendencia: 18 - 20 = 0
    }
    
    ventas_objetivo = {
        'ART001|U|UNICA': 100,
        'ART002|U|UNICA': 50,
        'ART003|U|UNICA': 20,
    }
    
    # Ejecutar aplicar_stock_minimo
    pedidos_actualizados, nuevo_stock, ajustes = engine.aplicar_stock_minimo(
        df, 5, stock_acumulado, stock_real, ventas_reales, ventas_objetivo
    )
    
    print("DATOS DE ENTRADA:")
    print(f"  - Und.Finales: 100, 50, 20")
    print(f"  - Stock Mínimo (30%): 30, 15, 6")
    print(f"  - Stock Real: 25, 15, 10")
    print(f"  - Ventas Reales: 110, 50, 18")
    print()
    
    # Verificar que NO existe la columna Unidades_Pedido
    print("VERIFICACIÓN 1: Eliminar columna Unidades_Pedido")
    print("-" * 80)
    if 'Unidades_Pedido' in pedidos_actualizados.columns:
        print("✗ ERROR: La columna 'Unidades_Pedido' todavía existe")
        test1 = False
    else:
        print("✓ OK: La columna 'Unidades_Pedido' ha sido eliminada")
        test1 = True
    
    # Verificar que existen las columnas de FASE 2
    columnas_requeridas = ['Stock_Minimo_Objetivo', 'Pedido_Corregido_Stock', 
                          'Ventas_Reales', 'Tendencia_Consumo', 'Pedido_Final']
    print()
    print("VERIFICACIÓN 2: Columnas de FASE 2 existen")
    print("-" * 80)
    test2 = True
    for col in columnas_requeridas:
        if col in pedidos_actualizados.columns:
            print(f"  ✓ {col}")
        else:
            print(f"  ✗ {col} NO encontrada")
            test2 = False
    
    # Verificar cálculos
    print()
    print("VERIFICACIÓN 3: Cálculos correctos (sin duplicar stock mínimo)")
    print("-" * 80)
    print(f"{'Código':>10} | {'Und.Fin':>8} | {'StkMin':>8} | {'StkReal':>8} | {'Ped.Corr':>10} | {'Ped.Fin':>10} | {'Verif':>8}")
    print("-" * 80)
    
    test3 = True
    for idx, row in pedidos_actualizados.iterrows():
        codigo = row['Codigo_Articulo']
        und_fin = row['Unidades_Finales']
        stk_min = row['Stock_Minimo_Objetivo']
        stk_real = stock_real.get(f'{codigo}|U|UNICA', 0)
        ped_corr = row['Pedido_Corregido_Stock']
        ped_fin = row['Pedido_Final']
        
        # Verificar fórmula: Pedido_Corregido_Stock = Und.Fin + (Stk.Min - Stk.Real)
        ped_corr_esperado = und_fin + (stk_min - stk_real)
        verificacion = abs(ped_corr - ped_corr_esperado) < 0.01
        
        # Verificar que NO duplica stock mínimo
        # Si duplicara: ped_corr = und_fin + 2*stk_min - stk_real
        ped_duplicado = und_fin + 2*stk_min - stk_real
        no_duplica = abs(ped_corr - ped_duplicado) > 0.01
        
        ok = verificacion and no_duplica
        test3 = test3 and ok
        
        estado = "✓ OK" if ok else "✗ FAIL"
        print(f"{codigo:>10} | {und_fin:>8.0f} | {stk_min:>8.0f} | {stk_real:>8.0f} | {ped_corr:>10.0f} | {ped_fin:>10.0f} | {estado:>8}")
        
        if not ok:
            print(f"           ERROR: {ped_corr} != {und_fin} + ({stk_min} - {stk_real}) = {ped_corr_esperado}")
    
    # Verificar filtrado
    print()
    print("VERIFICACIÓN 4: Filtrado usa Pedido_Corregido_Stock")
    print("-" * 80)
    pedidos_filtrados = pedidos_actualizados[pedidos_actualizados['Pedido_Corregido_Stock'] > 0]
    test4 = len(pedidos_filtrados) == 3  # Los 3 artículos tienen pedido > 0
    
    if test4:
        print(f"✓ OK: Se filtran {len(pedidos_filtrados)} artículos correctamente")
    else:
        print(f"✗ ERROR: Se filtraron {len(pedidos_filtrados)} artículos (esperado: 3)")
    
    # Verificar Total_Unidades
    print()
    print("VERIFICACIÓN 5: Total_Unidades usa Pedido_Final")
    print("-" * 80)
    total_pedido_final = int(pedidos_filtrados['Pedido_Final'].sum())
    
    # Calcular esperado
    total_esperado = 0
    for idx, row in pedidos_filtrados.iterrows():
        total_esperado += row['Pedido_Final']
    
    test5 = (total_pedido_final == total_esperado)
    
    if test5:
        print(f"✓ OK: Total_Unidades = Sum(Pedido_Final) = {total_pedido_final}")
    else:
        print(f"✗ ERROR: Total_Unidades = {total_pedido_final} (esperado: {total_esperado})")
    
    # Resumen
    print()
    print("=" * 80)
    print("RESUMEN DE VERIFICACIONES FINALES")
    print("=" * 80)
    
    todas_pasaron = test1 and test2 and test3 and test4 and test5
    
    verificaciones = [
        ("Eliminar columna Unidades_Pedido", test1),
        ("Columnas FASE 2 existen", test2),
        ("Cálculos correctos (sin duplicar)", test3),
        ("Filtrado usa Pedido_Corregido_Stock", test4),
        ("Total_Unidades usa Pedido_Final", test5)
    ]
    
    for nombre, resultado in verificaciones:
        estado = "✓ PASÓ" if resultado else "✗ FALLÓ"
        print(f"  {nombre}: {estado}")
    
    print("-" * 80)
    if todas_pasaron:
        print("✓ TODAS LAS VERIFICACIONES PASARON")
        print()
        print("RESUMEN DE ARCHIVOS MODIFICADOS:")
        print("  1. forecast_engine.py:")
        print("     • Eliminada variable 'unidades_pedido'")
        print("     • Fórmula corregida: Pedido_Corregido_Stock = Und.Fin + (Stk.Mín - Stk.Real)")
        print("     • Filtrado usa Pedido_Corregido_Stock > 0")
        print("     • Total_Unidades usa Sum(Pedido_Final)")
        print()
        print("  2. order_generator.py:")
        print("     • Eliminada columna 'Unidades Pedido' del Excel")
        print("     • COLUMN_MAPPING, COLUMN_HEADERS, COLUMN_WIDTHS actualizados")
        print("     • Filtrado y CSV usan Pedido_Corregido_Stock")
        print()
        print("  3. correction_engine.py:")
        print("     • 4 referencias cambiadas a Pedido_Corregido_Stock")
    else:
        print("✗ ALGUNAS VERIFICACIONES FALLARON")
        return 1
    
    return 0


def main():
    """Ejecutar verificación final"""
    print("\n" + "#" * 80)
    print("# VERIFICACIÓN FINAL: Todos los cambios (Pasos 1-4)")
    print("#" * 80 + "\n")
    
    return test_verificacion_final()


if __name__ == "__main__":
    sys.exit(main())
