# EDS v8.3 — optimización incremental y reproducible

Esta versión **no es una migración vectorizada integral**. Conserva todas las reglas, los JSON públicos, el histórico semanal/mensual, los tooltips y la auditoría privada de v8.2.

## Mejoras

- Iteración de registros con `itertuples` en lugar de `iterrows`, conservando el contrato `row.get()` de las reglas.
- Lectura opcional Polars de CSV y Parquet (`scan_csv` / `scan_parquet` con ejecución streaming), con adaptación a pandas para el motor original.
- En CSV, Polars conserva columnas como texto para proteger CLUES y claves con ceros iniciales.
- Los archivos Excel se siguen leyendo con pandas/openpyxl.

## Uso (Windows CMD)

```bat
pip install -r requirements_auditoria.txt
set EDS_MOTOR_LECTURA=polars
python scripts\calidad_5_modulos.py
```

Para usar el lector anterior: `set EDS_MOTOR_LECTURA=pandas`.

## Advertencias

- La conversión Polars→pandas implica uso adicional de memoria. Para CSV grandes, puede no ser mejor que pandas.
- Las reglas y el armado de listas de incumplimientos siguen en Python. Millones de incumplimientos pueden consumir mucha RAM.
- Las fechas, nulos y valores inferidos pueden diferir entre motores, especialmente en Parquet. Antes de publicar, comparar resultados de ambos motores usando el mismo conjunto de pruebas.
- Para una mejora sustancial será necesaria una **segunda etapa**: reglas vectorizadas en Polars, acumuladores por lotes y escritura Parquet incremental, sin cargar todos los incumplimientos en memoria.
- El visor privado escucha solo en 127.0.0.1. No suba `auditoria_privada` a GitHub.
