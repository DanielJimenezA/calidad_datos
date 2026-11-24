# Consolidación previa a v8.7
- Base: RAR entregado por el usuario; se recuperan cambios v8.5.2 y v8.6.
- Se eliminan duplicados MOC-82 y MOC-86 y se registran seis columnas nuevas MOCE.
- Se conserva el cálculo de cumplimiento por evaluaciones de v8.6.
- Se bloquea el procesamiento ante columnas faltantes para evitar publicar resultados incompletos como NO APLICA. La categoría NO EVALUADO aún no se implementa.
- Se corrigen condiciones nulas, longitud de causa directa y matrícula numérica.
- Los JSON incluidos son históricos anteriores: regenerar los cinco módulos antes de publicar.
- Pendientes: migración vectorizada a Polars, cobertura, Pareto, mapa de calor, integración auditoría y pruebas de producción.
