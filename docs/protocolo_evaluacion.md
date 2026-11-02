# Protocolo de evaluación

Define, **antes** de calcular los resultados, cómo se medirá si los niveles de inventario propuestos
(Objetivo 3) representan una mejora frente a la práctica actual del hospital. Responde al ajuste 2 de la
retroalimentación del ACA 1. Los umbrales son propuestas de los autores y deben validarse con el
servicio farmacéutico y con el director del proyecto.

## 1. Qué se evalúa

| Elemento | Definición |
|---|---|
| Unidad de análisis | Producto (`CODIGO_PROPIO`) para la evaluación; producto × servicio farmacéutico para los niveles finales |
| Horizonte | Mensual. Tiempo de reposición L = 0,5 mes y período de revisión R = 0,5 mes (H = L + R = 1 mes) |
| Propuesta | Mínimo (stock de seguridad), punto de reposición y máximo según patrón de frecuencia y celda ABC-VED |
| Datos | 29 meses (2024-01 a 2026-05) de la base analítica; ventana de estimación de 12 meses |

## 2. Líneas base

El sistema institucional **no tiene niveles mínimo y máximo definidos para la mayoría de los productos**:
en `Disponibilidad.xlsx`, el 97,84 % de las combinaciones producto-bodega figura con mínimo y máximo en 0
(`outputs/tablas/linea_base_inventario_actual.csv`). Por eso la línea base es el método empírico que
resulta de esa práctica, que se representa con dos reglas simples (y una de ablación):

| Código | Regla | Qué representa |
|---|---|---|
| B1 | Nivel máximo = consumo del último mes | Reponer lo que se consumió (decisión reactiva) |
| B2 | Nivel máximo = promedio de todo el histórico disponible, sin stock de seguridad | Promedio histórico simple (puede sobredimensionar si la demanda baja) |
| A1 (ablación) | Promedio de los últimos 12 meses, sin stock de seguridad ni reglas ABC-VED | Aísla el aporte del stock de seguridad y de la clasificación |
| B2k (control) | B2 multiplicada por un factor k para igualar el capital promedio de la propuesta | Permite comparar servicio **a capital equivalente** |

> Pendiente de validación: confirmar con la jefatura del servicio farmacéutico cómo se decide hoy cuánto
> comprar, para ajustar B1 y B2 si fuera necesario.

## 3. Métricas

Evaluación de origen móvil (*walk-forward*): para cada mes de origen *t* (desde el mes 12) se estiman los
parámetros **solo con datos hasta *t*** y se comparan con el consumo real del mes *t + 1*.

| Métrica | Fórmula | Lectura |
|---|---|---|
| Fill rate (unidades y valor) | Σ min(D, S) / Σ D | Proporción de la demanda que habría podido atenderse con el nivel máximo S |
| Disponibilidad ante demanda | Series con D > 0 y S ≥ D / series con D > 0 | Proporción de productos-mes sin faltante |
| Inventario promedio teórico (valor) | Σ costo × [S + máx(S − D, 0)] / 2 | Capital inmovilizado teórico por mes |
| Valor de demanda no cubierta | Σ costo × máx(D − S, 0) | Faltante expresado en pesos |

Cada métrica se reporta por criticidad (V, E, D) y global, con intervalo de confianza del 95 % por *bootstrap*
sobre los meses de origen y prueba de Wilcoxon pareada entre la propuesta y cada línea base.

## 4. Criterios de éxito (a priori)

| Criterio | Condición | Justificación |
|---|---|---|
| C1 | Disponibilidad ante demanda en productos vitales ≥ 95 % | Los vitales no deben faltar |
| C2 | Fill rate (valor) de la propuesta mayor que el de B1 y el de B2 | Mejora frente a la práctica actual |
| C3 | A capital equivalente, fill rate de la propuesta mayor que el de B2k | La mejora no se debe solo a mantener más inventario |

Se afirma que el Objetivo 3 se cumplió cuando se cumplen C1, C2 y C3. Si alguno no se cumple, se informa tal
cual y se analiza la causa (no se ajustan los umbrales después de ver los resultados).

El capital adicional que requiere la propuesta frente a B2 se informa como resultado, no como criterio:
la restricción presupuestal se trata en el Objetivo 4 (escenarios).

## 5. Control de fuga de información

- La clasificación ABC, el patrón de frecuencia y los parámetros se recalculan en cada origen con la
  ventana que termina en ese mes.
- Las **novedades** (descontinuado, a necesidad, Invima vencido, etc.) describen el estado a la fecha de
  corte del archivo; **no se aplican en la evaluación retrospectiva**, solo en los niveles finales.
- El costo unitario y la clasificación VEN son atributos vigentes de la Biblia (no tienen historia): se tratan
  como constantes y se declara como limitación (el valor de meses antiguos se expresa a costo vigente).

## 6. Sensibilidad

Ventana de estimación (6, 12, 24 meses), tiempo de reposición (L = 0,25; 0,5; 0,75 meses con H = 1) y
desplazamiento del nivel de servicio (−5, 0, +3 puntos porcentuales).

## 7. Limitaciones conocidas

- No hay datos de tiempos de reposición ni de faltantes reales por producto: L es un supuesto y el
  "faltante" es teórico (demanda mayor que el nivel máximo), no un quiebre observado.
- La demanda registrada es consumo (salidas), no demanda real: si hubo faltantes, la demanda queda subestimada.
- El inventario promedio teórico supone reposición a S al inicio de cada mes.
