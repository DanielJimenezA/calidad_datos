# v8.4 — Auditoría visible y reglas de horas delimitadas

- Se retira exclusivamente la regla **MOC-02 (FECHA_ATENCION)** del catálogo de calidad. Se conserva la lectura de `FECHA_ATENCION` como referencia temporal para agrupar el histórico MOCE por semana y mes; no se evalúa como incumplimiento. Si falta una fecha válida, el registro se excluye de las series temporales según la lógica existente.
- Las reglas de formato 24 horas se aplican exclusivamente a **IND_HORA_CITA**, **IND_INI_ATENCION** e **IND_FIN_ATENCION**; no se cambian las validaciones de fecha/hora de los otros módulos.
- Se añade **ARCHIVO_ORIGEN** a cada incidencia generada en nuevas ejecuciones.
- Se incorpora un acceso visible a la auditoría privada desde el final del tablero; la lista solo se presenta en el visor local, nunca en GitHub Pages.
- El visor privado muestra valor observado, archivo, fila y motivo; añade descarga CSV de los resultados filtrados (máximo 100,000 filas por descarga). El Parquet conserva todos los resultados.
- **No incluye aún migración vectorizada completa a Polars ni streaming del motor de reglas**. El procesamiento sigue siendo el de v8.3.2.

## Ejecución

```bat
pip install -r requirements_auditoria.txt
python scripts\calidad_5_modulos.py
python auditoria_local\app.py --carpeta "C:\ruta\de\salida\auditoria_privada"
```

Abre `http://127.0.0.1:8765/` en el equipo local. Reprocesa los archivos para que el catálogo y las cifras reflejen la eliminación de MOC-02 y para obtener ARCHIVO_ORIGEN en los Parquet. No publiques `auditoria_privada`.
