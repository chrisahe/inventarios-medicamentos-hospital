# Bitácora de decisiones

Registro de las decisiones tomadas durante la depuración y el modelado. Las decisiones D-01 a D-10 se
reconstruyen del ACA 1; D-11 en adelante corresponden al ACA 2. La columna **Estado** indica si la
decisión ya fue validada con el servicio farmacéutico o está pendiente.

| ID | Fase | Decisión | Alternativa considerada | Justificación / evidencia | Estado |
|---|---|---|---|---|---|
| D-01 | Datos | Retirar manualmente campos de pacientes, ventas y datos sensibles antes del procesamiento; conservar 7 variables | Procesar la base completa | Principio de mínima información; no se necesitan datos individuales | Aplicada |
| D-02 | Datos | Un producto en varias bodegas no es duplicado; se conserva producto-bodega-unidad funcional | Colapsar por producto | 1.908 códigos en más de una bodega (2024) | Aplicada |
| D-03 | Depuración | Excluir bodegas con `Se utiliza = No` en la base de Bodegas | Conservar todas | Criterio institucional | Aplicada |
| D-04 | Depuración | Excluir productos cuya descripción **inicia** con "NPT" (214.144 registros) | Excluir cualquier descripción que contenga "NPT" | Evita retirar productos que solo contienen esas letras | Aplicada |
| D-05 | Depuración | Bodega 829 (Nutriciones): revisión por producto, no exclusión general | Excluir toda la bodega | Observación específica en la fuente | Aplicada |
| D-06 | Depuración | Excluir 69 registros con cantidad ≤ 0 (68 negativos, 1 cero) | Conservarlos como devoluciones (plan inicial de la metodología) | La revisión de la fuente mostró que eran errores del sistema | Aplicada |
| D-07 | Depuración | Excluir oxígeno medicinal (código 100020287; 43.153 registros) | Conservarlo | Lo gestiona Biomédica, no el servicio farmacéutico; distorsionaba la escala | Aplicada |
| D-08 | Depuración | Conservar la bodega 855 ("Todas") como categoría histórica independiente | Excluirla o repartirla | Es consumo real; se declara como limitación de granularidad | Aplicada |
| D-09 | Fuentes | No usar Novedades como variable temporal del consumo | Fechar las novedades | Solo 4,5 % tiene fecha válida y ninguna cae en 2024-2026 | Aplicada |
| D-10 | EDA | Patrones de frecuencia descriptivos: sostenido 24-29, frecuente 12-23, intermitente 4-11, esporádico 1-3 meses | Otros umbrales | Caracterización del Objetivo 1 | Aplicada |
| D-11 | Costos | Eliminar 2.692 filas duplicadas exactas de la Biblia y 1 fila sin código | Conservarlas | Duplicados exactos inflan el conteo y no aportan información | Aplicada |
| D-12 | Costos | Costo unitario en cuatro niveles: (1) mediana de registros activos del producto en la misma bodega, (2) mediana de activos del producto, (3) mediana de inactivos, (4) sin costo | Promedio simple; último costo; costo de una sola bodega | 589 productos tienen más de un costo activo (mediana de la razón máx/mín 1,84); 2.425 productos solo tienen costo en registros inactivos; 408 no tienen costo | Pendiente: confirmar con finanzas si `COSTO FINAL` es costo unitario en COP |
| D-13 | ABC | Valor de consumo = cantidad × costo vigente (un solo precio por producto) | Valorar al costo de cada mes | La Biblia no tiene historia de precios; un precio único evita que cambios de precio alteren el ranking | Aplicada (limitación declarada) |
| D-14 | ABC | Ventana principal = últimos 12 meses; sensibilidad con 29 meses y por año | Usar los 29 meses | El consumo cayó en 2025-2026; la ventana reciente refleja la demanda vigente | Aplicada |
| D-15 | ABC | Cortes 80 % / 95 % (A / B / C); sensibilidad 70/90 y 75/95 | Otros cortes | Convención más usada; se reporta el efecto de cambiarlos | Aplicada |
| D-16 | VED | VED a partir de `Clasificacion VEN` de la Biblia; N (no esencial) se homologa a D (deseable) | Construir un criterio propio | Criterio clínico institucional ya existente; ningún producto tiene dos clasificaciones distintas | Pendiente: validar con el químico farmacéutico (la clasificación está sesgada hacia E: 79,5 % de los productos) |
| D-17 | VED | Producto sin clasificación → E (conservador), marcado para revisión | Dejarlo sin clase o asignarlo a D | Evita subestimar la criticidad | Aplicada |
| D-18 | VED | Novedad "Siempre deben haber existencias" → V | Ignorarla | Es una regla vigente de la institución (2 productos) | Aplicada |
| D-19 | Niveles | Mínimo = stock de seguridad; punto de reposición = demanda en L + stock de seguridad; máximo = demanda en L + R + margen de seguridad | Mínimo = demanda en L (fórmula simplificada del marco teórico del ACA 1) | Separa claramente los tres niveles pedidos en el Objetivo 3 y alinea con la política de revisión periódica | Aplicada |
| D-20 | Niveles | Nivel de servicio por celda ABC-VED (V: 98-99 %; E: 93-97 %; D: 85-92 %) | Un solo nivel para todos | Más servicio donde la criticidad es mayor; menos donde el producto es costoso y no crítico | Pendiente: validar con el servicio farmacéutico |
| D-21 | Niveles | Tiempo de reposición L = 0,5 mes y revisión R = 0,5 mes | Tiempos por proveedor | No hay datos de tiempos de entrega; se hace sensibilidad con L = 0,25 y 0,75 | Pendiente: confirmar con compras |
| D-22 | Niveles | Sostenido/frecuente → modelo normal; intermitente → cuantil empírico del consumo mensual; esporádico → solo vitales con existencia (máximo mensual de la ventana), el resto bajo demanda | Un solo modelo para todos | El 62,99 % de los productos es esporádico y el supuesto normal no aplica con tantos meses en cero | Aplicada |
| D-23 | Niveles | Estimación sobre ventana móvil de 12 meses | Promedio histórico completo | Tendencia a la baja del consumo (Tabla 19 del ACA 1) | Aplicada |
| D-24 | Niveles | Niveles finales por producto × servicio farmacéutico; la evaluación retrospectiva se hace por producto | Solo por producto | El objetivo general pide niveles por servicio; a nivel producto-servicio la demanda es más intermitente | Aplicada. "Almacén" se trata como un servicio más; confirmar si sus salidas duplican las de los servicios satélite |
| D-25 | Evaluación | Las novedades solo se aplican a los niveles finales, no al *backtest* | Aplicarlas también al *backtest* | Representan el estado vigente: aplicarlas hacia atrás sería fuga de información | Aplicada |
| D-26 | Evaluación | Línea base B1 (último mes) y B2 (promedio histórico) más ablación A1 y control a capital equivalente B2k | Comparar contra los parámetros del sistema | El 97,84 % de las combinaciones producto-bodega tiene mínimo y máximo en 0 | Pendiente: confirmar con la jefatura cómo se compra hoy |
| D-27 | Datos | Unir por `CODIGO_PROPIO` (no por descripción) | Unir por descripción | 802 códigos de la Biblia tienen más de una descripción | Aplicada |
| D-28 | Evaluación | El criterio C1 (vitales ≥ 95 %) no se cumple (90,06 %); no se modifica el umbral y se documenta la causa (vitales esporádicos sin consumo previo) | Ajustar el umbral o la ventana tras ver el resultado | El protocolo define los umbrales a priori; `outputs/tablas/o3_diagnostico_vitales.csv` | Aplicada. Evaluar existencia mínima institucional en el Objetivo 4 |
| D-29 | Datos | Comparar el consumo por servicio entre años con cautela: Almacén no registra consumo en ene–may 2024 y concentra 4,2 millones de unidades en ene–may 2025 | Interpretar la caída de Hospitalización como disminución real | Hospitalización + Almacén baja −16,6 %, frente a −72,5 % de Hospitalización sola | Pendiente: confirmar con el servicio farmacéutico y con sistemas el cambio de registro (también oct–nov 2024 y dic 2025) |

## Hallazgos de revisión del documento del ACA 1

| Hallazgo | Acción |
|---|---|
| Referencia 4 (Dube et al., 2023) no corresponde al artículo citado: el artículo *Modelling and Analysis of Hospital Inventory Policies during COVID-19 Pandemic* (Processes 11(4), 1062) es de Rehman, Mian, Usmani, Abidi y Mohammed; el DOI correcto es 10.3390/pr11041062 | Corregir autores, DOI y el párrafo de antecedentes (el estudio compara políticas de punto de reorden, pedido periódico y *just-in-time* mediante simulación) |
| Referencia 5 (Muhoza et al., 2023) no corresponde al artículo: el estudio sobre medicamentos obstétricos en hospitales de la Provincia Sur de Ruanda es de Kabera y Mukanyangezi, BMC Health Services Research, 2024; DOI 10.1186/s12913-023-10459-x | Corregir autores, año, volumen y número de artículo y revisar el párrafo de antecedentes |
| Citas numéricas "(1,3)" en el texto | Convertir a citas autor-año (APA 7) con `scripts/convertir_citas_apa7.py` |
| Referencias sin sangría francesa ni DOI verificable | Aplicar sangría francesa en Word y verificar cada DOI |
| Texto duplicado en Alcance/Limitaciones y marcador "Tabla X" | Revisar en el documento fuente |
