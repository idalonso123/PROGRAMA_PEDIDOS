# Sistema de Pedidos de Compra - Vivero Aranjuez V2

Sistema modular para la generación automática de pedidos de compra con generación semanal programada y persistencia de estado.

## Estructura del Proyecto

```
vivero_pedidos_v2/
├── config/
│   └── config.json          # Configuración principal del sistema
├── data/
│   ├── state.json           # Estado persistente entre ejecuciones
│   ├── input/               # Archivos de entrada (Excel)
│   └── output/              # Archivos de salida generados
├── src/
│   ├── data_loader.py       # Carga y normalización de datos
│   ├── state_manager.py     # Persistencia de estado
│   ├── forecast_engine.py   # Motor de cálculo de pedidos
│   ├── order_generator.py   # Generación de archivos Excel
│   ├── scheduler_service.py # Control de ejecución programada
│   └── main.py              # Script principal de orchestración
├── logs/
│   └── sistema.log          # Logs de ejecución
├── requirements.txt         # Dependencias de Python
└── README.md                # Este archivo
```

## Instalación

1. Crear un entorno virtual (recomendado):
```bash
python -m venv venv
source venv/bin/activate  # En Linux/Mac
# o
venv\Scripts\activate  # En Windows
```

2. Instalar las dependencias:
```bash
pip install -r requirements.txt
```

## Configuración

El archivo `config/config.json` contiene todos los parámetros del sistema:

- **secciones**: Objetivos de venta semanales por sección
- **parametros**: Factores de crecimiento, stock mínimo, pesos ABC
- **festivos**: Incrementos por períodos especiales
- **secciones_activas**: Lista de secciones a procesar
- **horario_ejecucion**: Día y hora de ejecución programada
- **rutas**: Directorios de entrada, salida y estado

## Uso

### Ejecución Normal (Programada)

```bash
python main.py
```

El sistema se ejecutará automáticamente los domingos a las 15:00 horas según lo configurado en `config.json`.

### Modo Forzado (para pruebas)

Procesar una semana específica:

```bash
python main.py --semana 15
```

### Modo Continuo

Espera hasta el horario de ejecución programado:

```bash
python main.py --continuo
```

### Mostrar Estado

Ver el estado actual del sistema:

```bash
python main.py --status
```

### Resetear Estado

Borrar todo el historial y comenzar desde cero:

```bash
python main.py --reset
```

### Modo Verboso

Activar logging detallado para depuración:

```bash
python main.py --verbose
```

## Flujo de Ejecución

1. **Verificación de horario**: Comprueba si es el momento de ejecutar (domingo 15:00)
2. **Carga de configuración**: Lee `config/config.json`
3. **Carga de estado**: Lee `data/state.json` para obtener información previa
4. **Determinación de semana**: Calcula qué semana debe procesar
5. **Lectura de datos**: Carga archivos de Ventas, Costes y Clasificación ABC
6. **Cálculo de pedidos**: Aplica metodología de escalado y factores
7. **Generación de archivos**: Crea Excel con los pedidos calculados
8. **Guardado de estado**: Actualiza `data/state.json` con los resultados

## Archivos de Entrada

Los siguientes archivos deben estar en el directorio `data/input/`:

- **SPA_ventas.xlsx**: Histórico de ventas por artículo
- **SPA_coste.xlsx**: Costes, precios y proveedores
- **CLASIFICACION_ABC+D_*.xlsx**: Clasificación y acciones por categoría

## Archivos de Salida

Los archivos generados se guardan en `data/output/`:

- **Pedido_Semana_XX_YYYY-MM-DD.xlsx**: Pedido de la semana
- **Pedido_Semana_XX_YYYY-MM-DD_SECCION.xlsx**: Pedido por sección
- **Resumen_Pedidos_SECCION_YYYY-MM-DD.xlsx**: Resumen consolidado

## Programación Automática (cron/Linux)

Para ejecutar automáticamente cada domingo a las 15:00:

```bash
crontab -e
```

Añadir la línea:

```cron
0 15 * * 0 /ruta/a/venv/bin/python /ruta/a/vivero_pedidos_v2/main.py >> /ruta/a/vivero_pedidos_v2/logs/cron.log 2>&1
```

## Programación Automática (Windows)

Usar el Programador de Tareas de Windows:

1. Abrir Programador de Tareas
2. Crear tarea básica
3. Configurar para cada domingo a las 15:00
4. Acción: Iniciar programa
5. Programa/script: `python`
6. Argumentos: `main.py`
7. Iniciar en: `C:\ruta\a\vivero_pedidos_v2\`

## Estado del Sistema

El archivo `data/state.json` mantiene:

- **informacion_sistema**: Versión, última ejecución, última semana procesada
- **stock_acumulado**: Stock actual por artículo
- **historico_ejecuciones**: Registro de todas las ejecuciones
- **pedidos_generados**: Historial de archivos generados
- **metricas**: Estadísticas acumuladas

## Licencia

Sistema interno desarrollado para Vivero Aranjuez.

## Autor

Sistema de Pedidos Vivero V2
