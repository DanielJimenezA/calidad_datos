# Calidad de datos institucional — versión 6.0

## Ejecución

1. Coloque el catálogo en `tablero/data/catalogos/unidades.json` (incluido).
2. Ejecute `python scripts/calidad_5_modulos.py` y seleccione los archivos fuente.
3. El script genera el Excel privado de auditoría y `tablero/data/calidad.json` más `tablero/data/tendencia.json`.
4. Desde `tablero/`, ejecute `python -m http.server 8000` y visite http://localhost:8000.
5. Para GitHub Pages, copie el contenido de `tablero/` a la raíz o a `docs/` del repositorio; configure Pages con esa carpeta.

## Definiciones

- Registros con incumplimiento: registros únicos que fallan una o más reglas.
- Evaluaciones fallidas: suma de NO CUMPLE; un registro puede fallar varias reglas.
- EVALUABLES = CUMPLE + NO_CUMPLE; NO_APLICA queda fuera del denominador.
- Histórico: se guarda una observación por semana ISO, con métricas filtrables por módulo, entidad y regla.
- Si se selecciona una regla, el KPI de registros únicos no se presenta como exacto: aparece «—» porque el JSON público no contiene registros individuales. La gráfica de esa regla sí usa evaluaciones.

## Privacidad y alcance

El JSON público contiene sólo agregados por módulo, entidad, regla y unidad. Los detalles de paciente y valores observados quedan en el Excel privado. Antes de publicar, la institución debe autorizar incluso los agregados por unidad y revisar celdas pequeñas.

Esta entrega incluye análisis territorial por entidad y tabla de unidades, pero **no incluye todavía un mapa geográfico**; la incorporación de GeoJSON validado queda como siguiente mejora. El catálogo es de referencia, no se inventan nombres de unidades.

Los datos JSON iniciales son marcadores vacíos, no resultados reales. Ejecute Python para generar un corte.


## Comparador de semanas (actualización)

El tablero incluye dos selectores: **Semana de referencia** y **Semana a comparar**. Permite seleccionar cualquier par de semanas presentes en `tablero/data/tendencia.json`, además de un acceso rápido a la semana anterior. La gráfica muestra todas las semanas disponibles y la tarjeta de variación corresponde exclusivamente al par seleccionado. Se preservan los filtros de módulo, entidad y regla, las tablas y el resto de visualizaciones.

La variación se calcula en puntos porcentuales. Para calidad de registros, un aumento es favorable; para incumplimiento de una regla, una disminución es favorable. No se calcula variación si faltan datos para los filtros o cambia `reglas_version`. Se conserva una sola entrada por semana ISO. **Para disponer de semanas históricas reales**, procesa los cinco archivos de cada semana con la misma versión de reglas y conserva los cortes; el script actual usa la fecha de ejecución como fecha de corte y reemplaza una ejecución anterior de la misma semana. No basta con copiar el `calidad.json` v5 al histórico v6.

## Procesamiento doble: corte anterior y actual

Esta edición del generador solicita **dos carpetas**, cada una con cinco archivos (Cirugías, MOCE, Urgencias, Egresos UCI y Egresos), y la **fecha de corte** de cada carpeta en formato `AAAA-MM-DD`. La fecha se escribe en la consola o celda de Jupyter donde se ejecuta el script; la selección de carpetas se hace con el explorador de archivos. Las fechas deben pertenecer a semanas ISO diferentes.

1. Guarda los cinco archivos anteriores en una carpeta y los cinco actuales en otra.
2. Ejecuta `python scripts/calidad_5_modulos.py` (o `%run scripts/calidad_5_modulos.py` desde Jupyter).
3. Selecciona la carpeta anterior e introduce su fecha de corte.
4. Selecciona la carpeta actual e introduce su fecha de corte.
5. Elige dos nombres diferentes para los Excel privados de auditoría.
6. El script procesa ambos conjuntos con las mismas reglas v6 y guarda ambos cortes en `tablero/data/tendencia.json`. `tablero/data/calidad.json` corresponde exclusivamente al corte actual.
7. Abre el tablero y selecciona ambas semanas en el comparador.

**Interpretación:** Si las bases son acumulativas, la comparación representa la calidad del universo acumulado en dos fotografías temporales. No equivale a medir exclusivamente los nuevos registros ni permite identificar correcciones de registros antiguos sin un cruce de identificadores estables. Una fecha de corte introducida manualmente etiqueta el conjunto; el script no elimina automáticamente registros posteriores a esa fecha. Asegúrate de que cada carpeta realmente representa el corte indicado.

**Seguridad:** el Excel detallado debe permanecer privado. El tablero usa sólo agregados. Antes de publicar resultados institucionales por unidad, revisa autorizaciones y riesgo de celdas pequeñas.


## Comparación histórica (actualización)

El botón **Primera vs. última semana** elige automáticamente los extremos disponibles del histórico. También puedes elegir manualmente una semana de referencia y una posterior. La gráfica muestra los cortes intermedios disponibles en ese rango, sin inventar semanas faltantes. Para incorporar nuevas semanas conserva el archivo `tablero/data/tendencia.json` entre ejecuciones. La comparación exige la misma `reglas_version`; la evolución del acumulado no demuestra correcciones de registros históricos individuales.
