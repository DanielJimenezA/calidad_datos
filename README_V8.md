# EDS · Calidad de datos institucionales — v8.0 (entrega incremental)

## Componentes
- `index.html`, `css/`, `js/`, `data/`: tablero público con agregados; desplegar en GitHub Pages desde la raíz.
- `scripts/calidad_5_modulos.py`: motor original más reglas MOCE MOC-80 a MOC-86.
- `auditoria_local/app.py`: visor privado que consulta el Excel consolidado generado por el motor.

## Uso
1. Instalar las dependencias originales del motor (pandas, numpy, xlsxwriter, openpyxl, pyarrow según formatos de entrada).
2. Ejecutar `python scripts/calidad_5_modulos.py` y seleccionar los cinco archivos y un Excel de salida **fuera del repositorio público**, preferiblemente en `salidas_privadas/`.
3. Ejecutar `python auditoria_local/app.py --excel "RUTA_AL_EXCEL_CONSOLIDADO.xlsx"`.
4. Abrir `http://127.0.0.1:8765` desde el mismo equipo.
5. Para probar el tablero estático: `python -m http.server 8000` desde la raíz y abrir `http://127.0.0.1:8000`.

## Advertencias y pendientes
- No incluir Excel de auditoría ni archivos clínicos en el repositorio. `.gitignore` excluye `salidas_privadas/` y los `.xlsx`, pero revisar siempre `git status` antes de publicar.
- El visor local lee las hojas `*_Resultados` del Excel, sin copiar datos al directorio público.
- Las nuevas reglas de matrícula y servicio validan obligatoriedad, caracteres y longitud; las relaciones con nombres del médico y tipo de médico requieren catálogos oficiales, no disponibles en esta entrega.
- El orden horario asume mismo día; los cruces de medianoche se marcan como incumplimiento hasta definir criterio institucional.
- El histórico publicado corresponde al conjunto anterior de reglas. Reprocesar las cinco fuentes antes de interpretar series con las nuevas reglas. No mezclar versiones de reglas en una misma comparación.
- Esta entrega no incluye aún mapa de calor ni Pareto, ni exportación desde el visor. Se conservan las gráficas existentes.
