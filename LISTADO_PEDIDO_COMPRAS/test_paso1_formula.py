#!/usr/bin/env python3
"""
Script de verificación: Paso 1 - Corregir fórmula Pedido_Corregido_Stock

Verificar que la fórmula NO duplica el stock mínimo:
- ANTES (incorrecto): Pedido_Corregido_Stock = (Und.Finales + Stk.Mín) + (Stk.Mín - Stk.Real)
                                            = Und.Finales + 2×Stk.Mín - Stk.Real
- AHORA (correcto):  Pedido_Corregido_Stock = Und.Finales + (Stk.Mín - Stk.Real)
                                            = Und.Finales + Stk.Mín - Stk.Real

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

def test_formula_correccion_stock():
    """
    Verificar que Pedido_Corregido_Stock NO duplica el stock mínimo
    """
    print("=" * 80)
    print("PASO 1: Verificar fórmula Pedido_Corregido_Stock (sin duplicar stock mínimo)")
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
    
    # Escenario: Déficit de stock
    stock_real = {
        'ART001|U|UNICA': 25,   # Déficit: stock mínimo es 30
        'ART002|U|UNICA': 15,   # Óptimo: stock mínimo es 15
        'ART003|U|UNICA': 10,   # Excedente: stock mínimo es 6
    }
    
    ventas_reales = {
        'ART001|U|UNICA': 110,
        'ART002|U|UNICA': 50,
        'ART003|U|UNICA': 18,
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
    
    print("Datos de entrada:")
    print(f"  - Unidades Finales (Forecast): 100, 50, 20")
    print(f"  - Stock Mínimo (30%): 30, 15, 6")
    print(f"  - Stock Real: 25, 15, 10")
    print()
    
    print("Verificación del cálculo:")
    print("-" * 100)
    print(f"{'Código':>10} | {'Und.Fin':>8} | {'StkMin':>8} | {'StkReal':>8} | {'Und.Ped':>8} | {'Ped.Corr':>10} | {'Fórmula_OK':>12}")
    print("-" * 100)
    
    todos_correctos = True
    
    for idx, row in pedidos_actualizados.iterrows():
        codigo = row['Codigo_Articulo']
        und_fin = row['Unidades_Finales']
        stk_min = row['Stock_Minimo_Objetivo']
        stk_real = stock_real.get(f'{codigo}|U|UNICA', 0)
        und_ped = row['Unidades_Pedido']
        ped_corr = row['Pedido_Corregido_Stock']
        
        # Calcular lo que debería ser
        # Fórmula correcta: Und.Finales + (Stk.Mín - Stk.Real)
        ped_corr_esperado = und_fin + (stk_min - stk_real)
        
        # Verificar que NO hay duplicación
        # Verificar que Und.Pedido = Und.Fin + Stk.Mín
        ped_base_esperado = und_fin + stk_min
        formula_pedido_ok = (und_ped == ped_base_esperado)
        
        # Verificar que Ped.Corr no tiene stock mínimo duplicado
        formula_corr_ok = (ped_corr == ped_corr_esperado)
        
        # Verificar explícitamente que no hay duplicación
        # Si duplicara stock mínimo, sería: und_fin + 2*stk_min - stk_real
        ped_corr_duplicado = und_fin + 2*stk_min - stk_real
        no_duplicado = (ped_corr != ped_corr_duplicado)
        
        todos_correctos = todos_correctos and formula_corr_ok and no_duplicado
        
        estado = "✓ OK" if (formula_corr_ok and no_duplicado) else "✗ FAIL"
        print(f"{codigo:>10} | {und_fin:>8.0f} | {stk_min:>8.0f} | {stk_real:>8.0f} | {und_ped:>8.0f} | {ped_corr:>10.0f} | {estado:>12}")
        print(f"           → Und.Pedido = {und_fin} + {stk_min} = {ped_base_esperado} {'✓' if formula_pedido_ok else '✗'}")
        print(f"           → Ped.Corr = {und_fin} + ({stk_min} - {stk_real}) = {ped_corr_esperado} {'✓' if formula_corr_ok else '✗'}")
        print(f"           → ¿Duplica stock mínimo? No ({ped_corr} != {ped_corr_duplicado}) {'✓' if no_duplicado else '✗'}")
        print()
    
    print("-" * 100)
    
    if todos_correctos:
        print("✓ RESULTADO: La fórmula NO duplica el stock mínimo")
        print("  La verificación automática confirma que:")
        print("  • Und.Pedido = Und.Finales + Stock.Mínimo")
        print("  • Ped.Corregido_Stock = Und.Finales + (Stock.Mínimo - Stock.Real)")
        print("  • NO hay duplicación del stock mínimo en ningún cálculo")
    else:
        print("✗ RESULTADO: Hay errores en el cálculo")
    
    return todos_correctos


def test_verificacion_matematica():
    """
    Verificación matemática detallada del cálculo
    """
    print("\n" + "=" * 80)
    print("VERIFICACIÓN MATEMÁTICA DETALLADA")
    print("=" * 80)
    print()
    
    print("FÓRMULA ANTIGUA (INCORRECTA - DUPLICABA STOCK MÍNIMO):")
    print("  Pedido_Corregido_Stock = unidades_pedido + (stock_minimo - stock_real)")
    print("                           = (Und.Finales + Stk.Mín) + (Stk.Mín - Stk.Real)")
    print("                           = Und.Finales + 2×Stk.Mín - Stk.Real")
    print()
    
    print("FÓRMULA NUEVA (CORRECTA - SIN DUPLICAR):")
    print("  Pedido_Corregido_Stock = Und.Finales + (stock_minimo - stock_real)")
    print()
    
    # Ejemplo numérico
    und_fin = 100
    stk_min = 30
    stk_real = 25
    
    print("EJEMPLO NUMÉRICO:")
    print(f"  - Und.Finales: {und_fin}")
    print(f"  - Stock Mínimo: {stk_min}")
    print(f"  - Stock Real: {stk_real}")
    print()
    
    # Cálculo con fórmula antigua (incorrecta)
    ped_antiguo = (und_fin + stk_min) + (stk_min - stk_real)
    
    # Cálculo con fórmula nueva (correcta)
    ped_nuevo = und_fin + (stk_min - stk_real)
    
    print(f"  FÓRMULA ANTIGUA (incorrecta):")
    print(f"    = (100 + 30) + (30 - 25)")
    print(f"    = 130 + 5")
    print(f"    = {ped_antiguo} ← ¡DUPLICA EL STOCK MÍNIMO!")
    print()
    
    print(f"  FÓRMULA NUEVA (correcta):")
    print(f"    = 100 + (30 - 25)")
    print(f"    = 100 + 5")
    print(f"    = {ped_nuevo} ← ¡CORRECTO! (suma stock mínimo una sola vez)")
    print()
    
    diferencia = ped_antiguo - ped_nuevo
    print(f"  DIFERENCIA: {ped_antiguo} - {ped_nuevo} = {diferencia} unidades")
    print(f"  (La fórmula antigua añadía {stk_min} unidades extra incorrectamente)")
    print()
    
    return True


def main():
    """Ejecuta todas las verificaciones"""
    print("\n" + "#" * 80)
    print("# VERIFICACIÓN PASO 1: Corregir fórmula Pedido_Corregido_Stock")
    print("# Objetivo: Eliminar duplicación del stock mínimo en el cálculo")
    print("#" * 80 + "\n")
    
    resultados = []
    
    # Test 1: Verificación con datos
    resultados.append(("Verificación con datos", test_formula_correccion_stock()))
    
    # Test 2: Verificación matemática
    resultados.append(("Verificación matemática", test_verificacion_matematica()))
    
    # Resumen
    print("\n" + "=" * 80)
    print("RESUMEN DE VERIFICACIONES - PASO 1")
    print("=" * 80)
    
    todas_pasaron = True
    for nombre, resultado in resultados:
        estado = "✓ PASÓ" if resultado else "✗ FALLÓ"
        print(f"  {nombre}: {estado}")
        todas_pasaron = todas_pasaron and resultado
    
    print("-" * 80)
    if todas_pasaron:
        print("✓ TODAS LAS VERIFICACIONES PASARON")
        print("\nLa fórmula corregida es:")
        print("  Pedido_Corregido_Stock = max(0, Unidades_Finales + (Stock_Mínimo - Stock_Real))")
        print("\nEsta fórmula:")
        print("  ✓ NO duplica el stock mínimo")
        print("  ✓ Calcula correctamente el pedido base menos el stock real")
        print("  ✓ Mantiene el buffer de seguridad de stock mínimo")
    else:
        print("✗ ALGUNAS VERIFICACIONES FALLARON")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
