# Tablero de Calidad de Datos v7.0 — IMSS-BIENESTAR

Proyecto basado en la versión v6. **Las reglas originales de calidad permanecen sin cambios.**

## Importante sobre los JSON incluidos

Los JSON incluidos son **únicamente una demostración** generada a partir de una muestra anonimizada (14 filas por módulo aproximadamente). No representan cifras institucionales, y no deben publicarse como resultados reales. Los archivos históricos se regeneran al ejecutar el análisis sobre los cinco insumos acumulativos reales.

## Fechas clínicas para las semanas ISO

- MOCE: `FECHA_ATENCION`
- Cirugías: `Fecha de Realización` (alias normalizado internamente `Fecha de realización`)
- Urgencias: `fecha_hora_ingreso`
- Egresos UCI: `fecha_hora_ingreso`
- Egresos: `fecha_hora_egreso`

Las fechas son **criterios de agrupación**, no reglas nuevas de calidad. Fechas vacías o inválidas se excluyen de la serie semanal y se contabilizan en `registros_sin_fecha`.

## Ejecución

1. Instalar dependencias: `python -m pip install -r requirements.txt`.
2. Ejecutar `python scripts/calidad_5_modulos.py` desde la carpeta raíz del proyecto. Seleccionar los cinco archivos acumulativos solicitados por el explorador y la ruta de salida del Excel.
3. El proceso genera `tablero/data/calidad.json`, `tablero/data/tendencia.json` y `tablero/data/historico_semanal.json` y el Excel de auditoría.
4. Ejecutar `cd tablero` y `python -m http.server 8000`.
5. Abrir `http://localhost:8000`.
6. Para GitHub Pages publicar la carpeta `tablero` como raíz del sitio (o copiar su contenido a la carpeta publicada).

## Filtros y métricas

Los filtros de módulo, entidad, CLUES/nombre de unidad, regla, semana inicial/final y modo semanal/acumulado afectan KPI, gráficas, tablas y CSV. La comparación usa las semanas seleccionadas; los demás paneles muestran la semana final seleccionada. El acumulado se limita al año ISO de la semana final.

- Todas las reglas: `registros con incumplimiento / registros analizados` (registros únicos dentro de cada módulo).
- Regla específica: `NO_CUMPLE / EVALUABLES`, con `EVALUABLES = CUMPLE + NO_CUMPLE`.
- La variación de calidad se expresa en puntos porcentuales (calidad = 100 - % incumplimiento).
- Semanas sin registros aparecen sin dato, no como 0% de calidad.
- El selector comienza en S01 de 2026 y termina en la última semana ISO completa según la fecha local del navegador.

## Limitaciones y seguridad

- No se incluyen datos de pacientes individuales en los JSON; el Excel detallado de auditoría debe mantenerse privado.
- Los conteos de registros representan filas de origen (no deduplicación por CURP ni clave de episodio), tal como lo hace el motor v6.
- No se reescribieron las reglas originales. Cualquier cambio de reglas requiere una versión explícita y revisión institucional.
- Validar con los cinco acumulativos completos antes de publicar: la muestra anonimizada no garantiza todas las semanas, CLUES o casuísticas.
- Si una CLUES no figura en `catalogos/unidades.json`, se mostrará como nombre no disponible pero el filtro por clave seguirá funcionando.

## Cambios v7.1 (2026-10-20)

- Filtro de unidades explícito con casillas CLUES · nombre de unidad. Se pueden marcar varias unidades. El buscador localiza opciones, pero **no aplica filtros por sí solo**: debe marcarse una o varias casillas. «Todas las unidades» quita la selección y muestra todas.
- Los KPI globales **no cambian de definición** al elegir una regla: registros, registros únicos afectados, cumplimiento global y evaluaciones incumplidas se mantienen; la regla seleccionada tiene un KPI separado.
- Escala continua de colores pastel según **cumplimiento**: 0 % rojo, 50 % amarillo, 100 % verde. Sin evaluables se muestra gris.
- Las semanas se muestran desde S01 hasta la última semana 2026 presente en los agregados; los periodos sin registros muestran «—», no cero de cumplimiento. **La última semana presente no implica necesariamente que esté completa**: verificar el corte de fuentes.
- El histórico por CLUES permite sumar conteos de registros únicos por unidad y semana. Si los archivos de origen contienen duplicados clínicos, se consideran filas distintas, como en el motor v6.
- La regla seleccionada afecta gráficos y tabla de reglas; los KPI globales conservan su significado. La comparación histórica cambia a cumplimiento de la regla cuando ésta está seleccionada, y lo indica en la descripción.
- El proyecto incluye JSON de muestra anonimizada/demostración, no totales institucionales. Ejecutar el script con los archivos acumulativos reales para recalcularlos.

### Verificación rápida

1. `python -m http.server 8000` desde `tablero/` y abrir `http://localhost:8000`.
2. Elegir Cirugías y Chiapas; comprobar KPI de cumplimiento global y KPI de la regla CURP por separado.
3. Marcar una CLUES en el selector y verificar que cambian KPI, gráficas, tablas y CSV.
4. Cambiar semana final, modo semanal/acumulado y comparar denominadores.
5. Si un porcentaje aparece como 100 %, verificar numerador y denominador del JSON antes de interpretarlo como «todas las reglas fallan».


## Cambios v7.2
- Gráfica histórica con color continuo por cumplimiento en puntos y segmentos.
- Selector CLUES desplegable, buscable y multiselección, filtrado por entidad y módulo.
- Semanas del 2026 hasta la última semana ISO completa, sin proyectar semanas futuras.
- Diagnóstico de semanas sin registros o sin evaluaciones aplicables; no se interpolan valores inexistentes.
- No es necesario reprocesar Python si los JSON ya son correctos.
- **Nota**: Los JSON empaquetados en esta entrega conservan la marca `demo: true`; reemplácelos por sus JSON reales ya procesados antes de analizar indicadores institucionales.
