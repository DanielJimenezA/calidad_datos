# v8.1 — Exportación privada escalable

El detalle completo de incumplimientos se guarda en `auditoria_privada/*_Resultados.parquet`, junto al Excel seleccionado. **Nunca publicar esta carpeta.** El Excel mantiene sus hojas de resumen, reglas y una **muestra de hasta 10,000 filas por módulo**; la fuente íntegra es Parquet.

## Instalar

```bash
pip install -r requirements_auditoria.txt
```

## Ejecutar

1. Ejecuta `python scripts/calidad_5_modulos.py` y selecciona los cinco archivos de entrada y un Excel de salida.
2. Ejecuta `python auditoria_local/app.py --carpeta "C:\ruta\de\salida\auditoria_privada"`.
3. Abre `http://127.0.0.1:8765`.

El visor usa DuckDB para filtrar directamente los archivos Parquet sin cargar millones de filas en el navegador. Es local, no ofrece autenticación; protege la carpeta con permisos de Windows y no la subas a GitHub. Los JSON públicos siguen generándose tras la exportación Excel. Si el proceso no termina, revisa su salida antes de publicar.

**Pendiente**: exportación Excel de los resultados filtrados desde el visor, tooltips personalizados y gráficos Pareto/mapa de calor. Las hojas Excel de detalle son muestras, no resultados completos.
