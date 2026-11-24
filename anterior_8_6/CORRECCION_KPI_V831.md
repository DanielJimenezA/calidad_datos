# Corrección de alcance de KPI (v8.3.1)

Al seleccionar una regla, los KPI de registros afectados, tasa y evaluaciones incumplidas se calculan para esa regla. Con «Todas» se mantiene el cálculo global. Los indicadores y las gráficas de cumplimiento usan universos distintos: la tasa se divide entre registros analizados; el cumplimiento excluye NO APLICA.

El conteo de NO CUMPLE para una regla se interpreta como número de registros afectados siempre que cada regla se evalúe una vez por registro; esa es la estructura del histórico actual. No se deben sumar los incumplimientos de varias reglas para obtener registros únicos globales.

También se permite seleccionar reglas con cero incumplimientos. Los tooltips indican el alcance actual.

Prueba recomendada: seleccionar MOCE/MOC-01 y comprobar 33/20058 = 0.16%, mientras que cumplimiento es 20025/20058 = 99.84%.
