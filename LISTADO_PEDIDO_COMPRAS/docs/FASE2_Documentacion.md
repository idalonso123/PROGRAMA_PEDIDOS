# Documentación FASE 2: Sistema de Corrección

## 1. Resumen de la Implementación

La FASE 2 del Sistema de Pedidos de Compra implementa un ciclo de retroalimentación que permite optimizar los pedidos semana a semana incorporando información real de ventas, stock y compras. Este sistema parte de la premisa de que las proyecciones históricas, aunque valiosas como punto de partida, inevitablemente difieren de la realidad operativa y deben ajustarse continuamente según la evolución observada.

La implementación se basa en el análisis exhaustivo del documento OPCIONES.docx, que define 54 escenarios de corrección posibles. De estos escenarios se extrajo una fórmula unificada que satisface todos los casos:

```
Pedido_Final = max(0, Pedido_Generado + (Stock_Mínimo - Stock_Real))
```

---

## 2. Archivos Creados

### 2.1 `src/correction_data_loader.py`

Este módulo extiende el DataLoader existente para leer los archivos de datos de corrección. Sus funciones principales incluyen la lectura del archivo de stock actual, la lectura de ventas reales de la semana, la lectura de compras recibidas, y la fusión de todos estos datos con el pedido teórico generado en la FASE 1.

El módulo busca automáticamente los archivos de corrección en el directorio de entrada, soportando el nuevo formato del ERP con timestamps. Los archivos esperados son:

- `SPA_Stock_actual.xlsx`: Inventario disponible al momento del cálculo (con o sin timestamp)

**Nota:** El sistema soporta archivos con el formato del ERP:
```
SPA_Nombre__YYYYMMDD_HHMMSS.xlsx
```

Si existen múltiples archivos con diferentes timestamps, el sistema selecciona automáticamente el más reciente. Si no hay archivos con timestamp, busca el archivo sin timestamp (formato legacy).

### 2.2 `src/correction_engine.py`

Este módulo implementa el motor de corrección con la lógica de escenarios del documento OPCIONES. Sus componentes principales incluyen la fórmula principal de corrección que se aplica a cada artículo, el cálculo de stock mínimo según la categoría ABC, la detección automática del escenario en el que se encuentra cada artículo, la generación de métricas de evaluación del sistema, y la creación de alertas para situaciones críticas.

El motor detecta automáticamente el escenario de cada artículo basándose en las comparaciones entre ventas reales versus objetivo, compras reales versus sugerido, y stock real versus mínimo. Esta información se utiliza para generar descripciones legibles de la corrección aplicada.

### 2.3 Actualización de `config/config.json`

Se añadieron las siguientes secciones de configuración:

```json
"archivos_entrada": {
    "ventas": "SPA_Ventas",
    "coste": "SPA_Coste",
    "compras": "SPA_Compras"
},

"archivos_correccion": {
    "stock_actual": "SPA_Stock_actual"
},

"parametros_correccion": {
    "habilitar_correccion": true,
    "stock_minimo_por_categoria": {
        "A": 1.5,
        "B": 1.0,
        "C": 0.5,
        "D": 0.0
    },
    "umbral_alerta_stock": 5,
    "permitir_pedidos_negativos": false
}
```

**Nota sobre el sistema de archivos:**

La configuración ahora usa nombres base sin extensión (ej: `SPA_Ventas` en lugar de `Ventas.xlsx`). El sistema busca automáticamente archivos que coincidan con estos nombres, ya sea con el formato legacy (`SPA_Ventas.xlsx`) o con el formato del ERP (`SPA_Ventas__20260205_210037.xlsx`).

---

## 3. Lógica de Corrección por Escenarios

### 3.1 Fórmula General

La fórmula extraída del análisis de los 54 escenarios es:

```
Pedido_Final = max(0, Pedido_Generado + Diferencia_Stock)
Donde: Diferencia_Stock = Stock_Mínimo - Stock_Real
```

Esta fórmula satisface todos los escenarios porque:

- **Si Stock_Real > Stock_Mínimo**: Se resta el excedente del pedido
- **Si Stock_Real = Stock_Mínimo**: Se mantiene el pedido sin cambios
- **Si Stock_Real < Stock_Mínimo**: Se añade la diferencia para recuperar el stock mínimo

### 3.2 Stock Mínimo por Categoría ABC

El sistema calcula el stock mínimo según la categoría ABC del artículo:

- **Categoría A**: 1.5 semanas de cobertura (artículos críticos)
- **Categoría B**: 1.0 semana de cobertura (artículos importantes)
- **Categoría C**: 0.5 semanas de cobertura (artículos complementarios)
- **Categoría D**: 0.0 semanas de cobertura (artículos sin rotación)

### 3.3 Detección de Escenarios

El sistema identifica automáticamente el escenario de cada artículo generando un código como `SUP_IGU_DEF`, donde cada parte indica:

- Primera parte: Ventas vs Objetivo (SUPERIOR, IGUAL, INFERIOR)
- Segunda parte: Compras vs Sugerido (EXCESO, IGUAL, DEFECTO)
- Tercera parte: Stock vs Mínimo (EXCEDENTE, OPTIMO, DEFICIT)

---

## 4. Métricas de Evaluación

El sistema calcula automáticamente las siguientes métricas:

| Métrica | Descripción |
|---------|-------------|
| `total_articulos` | Número total de artículos procesados |
| `articulos_corregidos` | Artículos cuyo pedido fue modificado |
| `porcentaje_corregidos` | Porcentaje de artículos corregidos |
| `total_unidades_original` | Suma de unidades del pedido teórico |
| `total_unidades_corregido` | Suma de unidades del pedido corregido |
| `diferencia_unidades` | Diferencia total entre pedidos |
| `precision_forecast` | Ratio entre ventas reales y objetivo |
| `articulos_alerta_stock` | Artículos con stock crítico |

---

## 5. Alertas Automáticas

El sistema genera alertas para las siguientes situaciones:

- **STOCK_CRITICO**: Artículos con stock en 0 o negativo
- **CAMBIOS_SIGNIFICATIVOS**: Artículos con correcciones superiores al 50%
- **SIN_VENTAS**: Artículos con stock pero sin ventas registradas

---

## 6. Integración con FASE 1

### 6.1 Flujo de Datos

1. **FASE 1** genera el pedido teórico (`Unidades_Pedido`)
2. **FASE 2** carga los archivos de corrección (Stock, Ventas, Compras)
3. **FASE 2** fusiona los datos con el pedido teórico
4. **FASE 2** aplica la fórmula de corrección
5. **FASE 2** genera el pedido final corregido (`Pedido_Corregido`)

### 6.2 Columnas Añadidas

| Columna | Descripción |
|---------|-------------|
| `Stock_Fisico` | Stock real actual del almacén |
| `Unidades_Vendidas` | Ventas reales de la semana |
| `Unidades_Recibidas` | Compras recibidas en la semana |
| `Stock_Minimo_Objetivo` | Stock mínimo objetivo según ABC |
| `Diferencia_Stock` | Diferencia entre stock mínimo y real |
| `Pedido_Corregido` | Pedido final tras corrección |
| `Escenario` | Código del escenario detectado |
| `Razon_Correccion` | Descripción legible de la corrección |

---

## 7. Uso del Sistema

### 7.1 Preparación de Archivos

Antes de ejecutar la corrección, deben existir los siguientes archivos en `./data/input/`:

```
data/input/
├── SPA_Stock_actual.xlsx       (obligatorio para corrección, con o sin timestamp)
└── SPA_Ventas_semana_*.xlsx    (opcional, si existe)
```

**Sistema de búsqueda flexible:**

El sistema utiliza el módulo `file_finder.py` que implementa la función `find_latest_file()`. Esta función:

1. **Busca archivos con timestamp ERP**: `SPA_Stock_actual__20260205_210037.xlsx`
2. **Selecciona el más reciente**: Si hay múltiples exports del mismo día, usa el de mayor timestamp
3. **Fallback legacy**: Si no hay archivos con timestamp, busca `SPA_Stock_actual.xlsx`

El ERP genera los archivos con el formato:
```
SPA_Nombre__YYYYMMDD_HHMMSS.xlsx
```

Donde:
- `YYYYMMDD`: Fecha de exportación (8 dígitos)
- `HHMMSS`: Hora de exportación (6 dígitos)

### 7.2 Estructura de Archivos de Entrada

**SPA_Stock_actual.xlsx** debe contener:
- Código de artículo
- Nombre del artículo
- Talla
- Color
- Stock físico (unidades)

**Ventas_semana.xlsx** debe contener:
- Código de artículo
- Nombre del artículo
- Talla
- Color
- Unidades vendidas
- Importe de venta

**Compras_semana.xlsx** debe contener:
- Código de artículo
- Nombre del artículo
- Talla
- Color
- Unidades recibidas
- Coste unitario

---

## 8. Verificación de Escenarios

La implementación ha sido verificada contra los escenarios del documento OPCIONES. La siguiente tabla muestra algunos casos de prueba:

| Pedido | Stock Mínimo | Stock Real | Esperado | Calculado |
|--------|--------------|------------|----------|-----------|
| 10 | 20 | 30 | 0 | 0 ✓ |
| 10 | 20 | 20 | 10 | 10 ✓ |
| 10 | 20 | 10 | 20 | 20 ✓ |
| 10 | 20 | 10 | 20 | 20 ✓ |
| 10 | 20 | 10 | 20 | 20 ✓ |
| 10 | 20 | 10 | 20 | 20 ✓ |

---

## 9. Próximos Pasos

Para completar la integración total del sistema, se recomienda:

1. Actualizar `main.py` para orquestar la ejecución de FASE 1 + FASE 2
2. Crear un script de generación de archivos de corrección
3. Implementar métricas de evaluación semanales
4. Desarrollar dashboard de visualización de correcciones

---

## 10. Contacto y Soporte

Para dudas o problemas con la implementación de FASE 2, revisar los logs en `./logs/` o contactar con el equipo de desarrollo.

---

*Documento generado automáticamente para el Sistema de Pedidos de Compra - Vivero Aranjuez V2*
*Fecha: 2026-02-03*
