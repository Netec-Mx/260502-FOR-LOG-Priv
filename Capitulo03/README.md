# Práctica 13 — Construcción Manual de un Modelo de Promedio Móvil

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Crear (*Create*) |
| **Módulo** | 3 — Métodos Cuantitativos de Pronóstico |
| **Herramienta principal** | Microsoft Excel 2016 o superior |
| **Archivo de entrada** | `Lab13_Demanda_Semanal.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Lab13_PMS_[TuNombre].xlsx` |

---

## Descripción General

En esta práctica construirás **desde cero** dos modelos de Promedio Móvil Simple (PMS) — uno con ventana de 3 períodos (PMS-3) y otro con ventana de 5 períodos (PMS-5) — usando únicamente la función `PROMEDIO()` de Excel sobre rangos desplazados manualmente. Trabajarás con 24 semanas de demanda histórica de un producto de distribución, calcularás los errores de pronóstico para cada período disponible y crearás un gráfico de líneas comparativo con las tres series. Al finalizar, responderás preguntas de reflexión que conectan los resultados numéricos con decisiones operativas reales de inventario y picking.

> **Enfoque logístico:** A lo largo de la práctica, ten presente la pregunta central: *¿Cómo impacta el tamaño de ventana elegido en la capacidad de reaccionar ante cambios de demanda y en la planificación del reabastecimiento?*

---

## Objetivos de Aprendizaje

Al completar esta práctica serás capaz de:

- [ ] Construir manualmente las fórmulas de PMS-3 y PMS-5 en Excel usando `PROMEDIO()` con rangos desplazados, sin recurrir a funciones automáticas de pronóstico.
- [ ] Calcular el error simple de pronóstico (Demanda Real − Pronóstico) para cada período disponible en ambos modelos.
- [ ] Crear un gráfico de líneas con tres series (demanda real, PMS-3, PMS-5) e interpretar visualmente el efecto del tamaño de ventana sobre la suavización y reactividad del modelo.
- [ ] Justificar operativamente, en contexto logístico, cuándo conviene usar una ventana corta versus una ventana larga para productos con distintos perfiles de demanda.

---

## Prerrequisitos

### Conocimiento previo requerido

| Prerrequisito | Descripción |
|---|---|
| Contenido teórico completado | Lección 3.1 — Método de Promedio Móvil (fórmula, efecto de N, limitaciones) |
| Manejo de Excel | Ingreso de fórmulas, referencias relativas (`B2:B4`) y absolutas (`$B$2`), arrastre de fórmulas hacia abajo |
| Conceptos logísticos | Comprensión básica de reabastecimiento, nivel de servicio y picking |

### Acceso y archivos necesarios

| Recurso | Estado requerido |
|---|---|
| Microsoft Excel 2016 o superior instalado | ✅ Obligatorio |
| Archivo `Lab13_Demanda_Semanal.xlsx` | ✅ Descargado y guardado localmente |
| Al menos 500 MB de espacio libre en disco | ✅ Para guardar el archivo de trabajo |

---

## Entorno de Laboratorio

### Configuración de hardware recomendada

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Espacio en disco disponible | 500 MB | 2 GB |

### Software requerido

| Software | Versión mínima | Notas |
|---|---|---|
| Microsoft Excel | 2016 | Microsoft 365 recomendado |
| Sistema operativo | Windows 10 / macOS 11 | Excel para Mac es compatible |

### Preparación inicial del entorno

Antes de comenzar, ejecuta los siguientes pasos de configuración:

1. Abre el Explorador de archivos (Windows) o Finder (Mac) y localiza el archivo `Lab13_Demanda_Semanal.xlsx`.
2. Crea una **copia de trabajo** renombrando el archivo como `Lab13_PMS_[TuNombre].xlsx` antes de abrirlo. Esto preserva el archivo original.
3. Abre `Lab13_PMS_[TuNombre].xlsx` en Excel.
4. Verifica que el archivo contiene la hoja **"Datos"** con al menos las columnas **A** (Semana) y **B** (Demanda Real) con 24 filas de datos (S1–S24).
5. Confirma que Excel muestra los números correctamente (sin signos de error `#¿NOMBRE?` o celdas vacías inesperadas).

> **Nota para el instructor:** El archivo de práctica debe contener exactamente 24 semanas de demanda para un único SKU (por ejemplo, "Aceite Vegetal 1L"). Los valores deben oscilar en un rango realista (400–600 unidades/semana) con variabilidad moderada y sin estacionalidad marcada, para que ambos modelos produzcan pronósticos visualmente distinguibles.

---

## Desarrollo Paso a Paso

---

### Paso 1: Explorar y comprender la estructura del dataset

**Objetivo:** Familiarizarse con los datos de demanda semanal y preparar la hoja de trabajo para recibir los cálculos de pronóstico.

#### Instrucciones

1. Con el archivo `Lab13_PMS_[TuNombre].xlsx` abierto, haz clic en la hoja **"Datos"**.

2. Identifica y confirma la estructura de la tabla. Debe verse similar a esta:

   | Columna A | Columna B |
   |---|---|
   | Semana | Demanda Real (unidades) |
   | S1 | 480 |
   | S2 | 510 |
   | S3 | 495 |
   | … | … |
   | S24 | (valor) |

3. En la celda **C1**, escribe el encabezado:
   ```
   PMS-3 (Pronóstico)
   ```

4. En la celda **D1**, escribe el encabezado:
   ```
   PMS-5 (Pronóstico)
   ```

5. En la celda **E1**, escribe el encabezado:
   ```
   Error PMS-3
   ```

6. En la celda **F1**, escribe el encabezado:
   ```
   Error PMS-5
   ```

7. Aplica **negrita** a la fila 1 completa (encabezados). Selecciona A1:F1 y presiona `Ctrl + N` (Windows) o `Cmd + B` (Mac).

8. Guarda el archivo con `Ctrl + G` (Windows) o `Cmd + S` (Mac).

#### Resultado esperado

La hoja "Datos" debe mostrar la tabla original en columnas A y B, con cuatro nuevos encabezados en negrita en las columnas C, D, E y F. Las celdas C2:F25 deben estar vacías, listas para recibir los cálculos.

#### Verificación

- [ ] Los encabezados de C1 a F1 están presentes y en negrita.
- [ ] Las columnas A y B contienen exactamente 24 filas de datos (S1 a S24) sin celdas vacías ni errores.
- [ ] El archivo ha sido guardado con el nombre correcto (`Lab13_PMS_[TuNombre].xlsx`).

---

### Paso 2: Construir el modelo PMS-3 (ventana de 3 períodos)

**Objetivo:** Calcular manualmente el pronóstico de promedio móvil simple con N=3 para los períodos S4 a S24, comprendiendo por qué los primeros N−1 períodos no tienen pronóstico.

#### Instrucciones

1. Haz clic en la celda **C2** (fila correspondiente a S1). **Déjala vacía.** No hay pronóstico para S1 porque no existe historia previa suficiente.

2. Haz clic en la celda **C3** (fila correspondiente a S2). **Déjala vacía.** Tampoco hay pronóstico para S2 con N=3.

3. Haz clic en la celda **C4** (fila correspondiente a S3). **Déjala vacía.** Para calcular el pronóstico de S4, necesitamos los datos de S1, S2 y S3 — pero el pronóstico se coloca en el período *siguiente* al último período usado.

   > **Concepto clave:** El pronóstico de PMS-3 para el período t+1 se calcula usando los datos de los períodos t, t-1 y t-2. Por convención en esta práctica, el valor calculado se coloca en la misma fila del período que se está pronosticando. Es decir: la celda C5 contendrá el pronóstico para S4 (calculado con S1, S2, S3)... 
   
   > **Aclaración de alineación:** Para mantener la tabla interpretable, alinearemos el pronóstico con el período al que corresponde. El pronóstico para la semana 4 (S4) se coloca en C5 (fila de S4). Revisa la tabla de ejemplo del contenido teórico: el pronóstico de Abril usa Enero+Febrero+Marzo y se muestra en la fila de Abril. Seguiremos esa misma convención.

   **Corrección de alineación:** Dado que la fila 1 es de encabezados y los datos comienzan en la fila 2 (S1), la alineación correcta es:
   - Fila 2 → S1 (sin pronóstico)
   - Fila 3 → S2 (sin pronóstico)
   - Fila 4 → S3 (sin pronóstico)
   - Fila 5 → S4 → **primer pronóstico PMS-3** (promedio de S1, S2, S3)

4. Haz clic en la celda **C5** (fila de S4) e ingresa la siguiente fórmula:
   ```excel
   =PROMEDIO(B2:B4)
   ```
   Presiona `Enter`. Verás el promedio de las primeras tres semanas de demanda.

5. Haz clic nuevamente en **C5**. Ahora **arrastra la fórmula hacia abajo** hasta la celda **C25** (fila de S24):
   - Posiciona el cursor en la esquina inferior derecha de C5 hasta que aparezca una cruz negra (+).
   - Haz clic y arrastra hasta C25.
   - Suelta el mouse.

6. Verifica que las fórmulas se hayan desplazado correctamente haciendo clic en algunas celdas intermedias:
   - **C6** debe contener: `=PROMEDIO(B3:B5)` (pronóstico para S5, usando S2, S3, S4)
   - **C7** debe contener: `=PROMEDIO(B4:B6)` (pronóstico para S6, usando S3, S4, S5)
   - **C25** debe contener: `=PROMEDIO(B22:B24)` (pronóstico para S24, usando S21, S22, S23)

7. Las celdas **C2, C3 y C4** deben permanecer **vacías** (no escribas ningún valor ni fórmula en ellas).

#### Resultado esperado

La columna C debe mostrar:
- Celdas C2, C3, C4: **vacías** (sin valor)
- Celdas C5 a C25: **valores numéricos decimales** representando el promedio de las tres semanas anteriores para cada período

Ejemplo de valores esperados (pueden variar según tu dataset):

| Fila | Semana | Demanda Real | PMS-3 |
|---|---|---|---|
| 2 | S1 | 480 | *(vacío)* |
| 3 | S2 | 510 | *(vacío)* |
| 4 | S3 | 495 | *(vacío)* |
| 5 | S4 | 520 | **495.0** |
| 6 | S5 | 500 | **508.3** |
| 7 | S6 | 515 | **505.0** |

#### Verificación

- [ ] C5 contiene `=PROMEDIO(B2:B4)` y muestra un valor numérico.
- [ ] C25 contiene `=PROMEDIO(B22:B24)` y muestra un valor numérico.
- [ ] C2, C3 y C4 están vacías.
- [ ] No aparecen errores `#¡VALOR!`, `#¡REF!` ni `#¿NOMBRE?` en ninguna celda de la columna C.

---

### Paso 3: Construir el modelo PMS-5 (ventana de 5 períodos)

**Objetivo:** Aplicar la misma lógica de construcción manual para N=5, observando que ahora los primeros 4 períodos quedan sin pronóstico.

#### Instrucciones

1. Las celdas **D2, D3, D4, D5 y D6** deben permanecer **vacías**. Con N=5, los primeros 4 períodos no tienen pronóstico (se necesitan 5 datos previos para calcular el primer promedio).

2. Haz clic en la celda **D7** (fila de S6, que es el primer período que puede tener pronóstico PMS-5 usando S1 a S5) e ingresa:
   ```excel
   =PROMEDIO(B2:B6)
   ```
   Presiona `Enter`.

   > **¿Por qué D7 y no D6?** Siguiendo la misma convención de alineación: el pronóstico para S6 se calcula con los datos de S1, S2, S3, S4, S5. La fila de S6 es la fila 7 (dado que la fila 1 es encabezado). Por lo tanto, el primer pronóstico PMS-5 aparece en la fila 7.

3. Haz clic nuevamente en **D7** y **arrastra la fórmula hacia abajo** hasta la celda **D25** (fila de S24).

4. Verifica la alineación de las fórmulas:
   - **D7** debe contener: `=PROMEDIO(B2:B6)` (pronóstico para S6)
   - **D8** debe contener: `=PROMEDIO(B3:B7)` (pronóstico para S7)
   - **D25** debe contener: `=PROMEDIO(B20:B24)` (pronóstico para S24)

5. Las celdas **D2 a D6** deben permanecer **vacías**.

6. Guarda el archivo: `Ctrl + G` (Windows) o `Cmd + S` (Mac).

#### Resultado esperado

La columna D debe mostrar:
- Celdas D2 a D6: **vacías** (sin valor)
- Celdas D7 a D25: **valores numéricos** representando el promedio de las cinco semanas anteriores

Comparando con PMS-3, los valores de PMS-5 deben variar menos entre semanas consecutivas (más suavizados).

#### Verificación

- [ ] D7 contiene `=PROMEDIO(B2:B6)` y muestra un valor numérico.
- [ ] D25 contiene `=PROMEDIO(B20:B24)` y muestra un valor numérico.
- [ ] D2 a D6 están vacías.
- [ ] Los valores de la columna D son visiblemente más estables (menos variación entre filas) que los de la columna C.

---

### Paso 4: Calcular el error simple de pronóstico

**Objetivo:** Medir la diferencia entre la demanda real y cada pronóstico para evaluar qué tan cerca estuvo el modelo de la realidad en cada período.

#### Instrucciones

El **error simple** se define como:

```
Error = Demanda Real − Pronóstico
```

Un error positivo indica que el modelo subestimó la demanda (pronóstico menor a la realidad). Un error negativo indica sobreestimación.

**Para el Error PMS-3 (columna E):**

1. Las celdas **E2, E3 y E4** deben permanecer **vacías** (no hay pronóstico para esos períodos).

2. Haz clic en la celda **E5** e ingresa:
   ```excel
   =B5-C5
   ```
   Presiona `Enter`.

3. Arrastra la fórmula de **E5** hacia abajo hasta **E25**.

4. Verifica que **E5** calcule correctamente: si B5 = 520 y C5 = 495, entonces E5 debe mostrar **25**.

**Para el Error PMS-5 (columna F):**

5. Las celdas **F2 a F6** deben permanecer **vacías**.

6. Haz clic en la celda **F7** e ingresa:
   ```excel
   =B7-D7
   ```
   Presiona `Enter`.

7. Arrastra la fórmula de **F7** hacia abajo hasta **F25**.

**Formato de los errores:**

8. Selecciona el rango **E5:F25** y aplica formato de número con **1 decimal**:
   - Clic derecho → **Formato de celdas** → **Número** → Posiciones decimales: **1** → Aceptar.

9. Opcionalmente, aplica **formato condicional** para destacar errores grandes:
   - Selecciona E5:F25 → **Inicio** → **Formato condicional** → **Escala de color** → elige la escala verde-amarillo-rojo.
   - Esto permite identificar visualmente las semanas donde el modelo falló más.

10. Guarda el archivo.

#### Resultado esperado

Las columnas E y F deben mostrar valores positivos y negativos alternados, reflejando las semanas donde el pronóstico fue mayor o menor que la demanda real. Los errores de PMS-3 deben ser generalmente mayores en valor absoluto que los de PMS-5 en semanas de demanda estable, pero PMS-3 puede reaccionar mejor ante cambios bruscos.

#### Verificación

- [ ] E5 contiene `=B5-C5` y muestra un valor numérico (positivo o negativo).
- [ ] F7 contiene `=B7-D7` y muestra un valor numérico.
- [ ] No hay errores de fórmula en las columnas E y F.
- [ ] El formato de 1 decimal está aplicado correctamente en E5:F25.

---

### Paso 5: Crear el gráfico de líneas comparativo

**Objetivo:** Visualizar simultáneamente las tres series (demanda real, PMS-3 y PMS-5) para interpretar de forma gráfica el efecto del tamaño de ventana sobre la suavización y el rezago del pronóstico.

#### Instrucciones

**Selección de datos para el gráfico:**

1. Mantén presionada la tecla `Ctrl` y selecciona los siguientes rangos (selección múltiple no contigua):
   - **A1:A25** (columna Semana — eje X)
   - **B1:B25** (Demanda Real)
   - **C1:C25** (PMS-3)
   - **D1:D25** (PMS-5)

   > **Consejo:** Selecciona primero A1:D25 completo (es más fácil) y luego excluirás las columnas E y F del gráfico.

   **Método alternativo recomendado:** Selecciona el rango **A1:D25** completo (incluye las cuatro columnas necesarias).

**Insertar el gráfico:**

2. Ve a la pestaña **Insertar** en la cinta de opciones.

3. En el grupo **Gráficos**, haz clic en **Insertar gráfico de líneas o de área** (ícono con líneas).

4. Selecciona **Línea con marcadores** (la segunda opción en la primera fila).

5. Excel insertará un gráfico en la hoja. Haz clic en el gráfico para seleccionarlo y arrástralo debajo de la tabla de datos (aproximadamente desde la fila 27 hacia abajo).

**Configurar el eje horizontal (semanas):**

6. Haz clic derecho sobre el gráfico → **Seleccionar datos**.

7. En el panel derecho (**Etiquetas del eje horizontal**), haz clic en **Editar**.

8. Selecciona el rango **A2:A25** (las etiquetas S1 a S24) y haz clic en **Aceptar**.

9. Haz clic en **Aceptar** para cerrar el panel de selección de datos.

**Personalizar el gráfico:**

10. Haz doble clic en el **título del gráfico** y escribe:
    ```
    Comparativo de Pronósticos: Demanda Real vs PMS-3 vs PMS-5
    ```

11. Haz clic en la serie **Demanda Real** (línea azul por defecto) y en el panel de formato, cambia el color a **azul oscuro** y el grosor de línea a **2.25 pt** para destacarla.

12. Haz clic en la serie **PMS-3** y cambia el color a **naranja**.

13. Haz clic en la serie **PMS-5** y cambia el color a **verde**.

14. Asegúrate de que la **leyenda** sea visible (parte inferior del gráfico). Si no aparece: clic en el gráfico → ícono **+** (Elementos del gráfico) → activa **Leyenda → Inferior**.

15. Activa las **líneas de cuadrícula horizontales** para facilitar la lectura (ícono **+** → **Líneas de cuadrícula** → **Horizontales principales**).

16. Guarda el archivo.

#### Resultado esperado

El gráfico debe mostrar tres líneas claramente distinguibles:
- **Línea azul (Demanda Real):** la más irregular, con subidas y bajadas semanales.
- **Línea naranja (PMS-3):** sigue de cerca a la demanda real, con variaciones visibles pero suavizadas.
- **Línea verde (PMS-5):** la más suave de las tres, con menos picos y valles, pero con mayor rezago respecto a la demanda real.

Las líneas naranja y verde deben comenzar más tarde que la azul (C y D tienen celdas vacías al inicio), lo que es visualmente correcto.

#### Verificación

- [ ] El gráfico contiene exactamente tres series de datos con colores diferenciados.
- [ ] El eje horizontal muestra las etiquetas S1 a S24.
- [ ] La leyenda identifica correctamente cada serie.
- [ ] El título del gráfico está escrito correctamente.
- [ ] La línea PMS-5 es visualmente más suave (menos oscilante) que la línea PMS-3.

---

### Paso 6: Análisis estadístico rápido de los errores

**Objetivo:** Calcular estadísticas básicas de los errores para cuantificar el desempeño de cada modelo antes de la reflexión cualitativa.

#### Instrucciones

1. Haz clic en una celda vacía debajo del gráfico o en una nueva área de la hoja (por ejemplo, celda **A28**). Escribe el siguiente bloque de etiquetas y fórmulas:

   | Celda | Contenido a escribir |
   |---|---|
   | A28 | `Indicador` |
   | B28 | `PMS-3` |
   | C28 | `PMS-5` |
   | A29 | `Promedio del Error` |
   | A30 | `Desv. Estándar del Error` |
   | A31 | `Error Máximo (abs)` |
   | A32 | `Error Mínimo (abs)` |

2. En la celda **B29**, ingresa la fórmula para el promedio del error de PMS-3:
   ```excel
   =PROMEDIO(E5:E25)
   ```

3. En la celda **C29**, ingresa:
   ```excel
   =PROMEDIO(F7:F25)
   ```

4. En la celda **B30**, ingresa la desviación estándar del error PMS-3:
   ```excel
   =DESVEST(E5:E25)
   ```

5. En la celda **C30**, ingresa:
   ```excel
   =DESVEST(F7:F25)
   ```

6. En la celda **B31**, ingresa el error máximo en valor absoluto para PMS-3:
   ```excel
   =MAX(ABS(E5:E25))
   ```
   > **Importante:** Esta es una **fórmula matricial**. Después de escribirla, presiona `Ctrl + Shift + Enter` (Windows) o `Cmd + Shift + Enter` (Mac) en lugar de solo `Enter`. Excel mostrará llaves `{}` alrededor de la fórmula indicando que es matricial.

7. Repite el mismo proceso en **C31**:
   ```excel
   =MAX(ABS(F7:F25))
   ```
   Presiona `Ctrl + Shift + Enter`.

8. En **B32**, ingresa:
   ```excel
   =MIN(ABS(E5:E25))
   ```
   Presiona `Ctrl + Shift + Enter`.

9. En **C32**, ingresa:
   ```excel
   =MIN(ABS(F7:F25))
   ```
   Presiona `Ctrl + Shift + Enter`.

10. Aplica **negrita** a la fila 28 (encabezados de la tabla de indicadores).

11. Guarda el archivo.

#### Resultado esperado

Una tabla de resumen estadístico con 4 indicadores para cada modelo. Los valores típicos esperados (variarán según el dataset):

| Indicador | PMS-3 | PMS-5 |
|---|---|---|
| Promedio del Error | Cercano a 0 (puede ser positivo o negativo) | Cercano a 0 |
| Desv. Estándar del Error | Mayor que PMS-5 | Menor que PMS-3 |
| Error Máximo (abs) | Valor relativamente alto | Valor más moderado |
| Error Mínimo (abs) | Valor pequeño | Valor pequeño |

> **Interpretación esperada:** PMS-3 tendrá mayor variabilidad en sus errores (desviación estándar más alta) porque reacciona más a los cambios, pero también puede capturar mejor los picos reales de demanda. PMS-5 producirá errores más consistentes pero puede perder señales importantes.

#### Verificación

- [ ] Las cuatro fórmulas en B29:C32 muestran valores numéricos sin errores.
- [ ] Las fórmulas en B31:C32 y B32:C32 muestran llaves `{}` indicando que son fórmulas matriciales.
- [ ] La desviación estándar de PMS-5 es menor o igual a la de PMS-3 (esto es esperable con demanda moderadamente estable).

---

### Paso 7: Reflexión guiada — Implicaciones logísticas

**Objetivo:** Conectar los resultados cuantitativos del modelo con decisiones operativas reales de inventario, reabastecimiento y planificación de picking.

#### Instrucciones

1. Crea una nueva hoja en el archivo. Haz clic derecho sobre la pestaña **"Datos"** → **Insertar** → **Hoja** → **Aceptar**. Renombra la nueva hoja como **"Reflexión"** haciendo doble clic sobre su pestaña.

2. En la hoja **"Reflexión"**, escribe las siguientes preguntas en la columna A (una por celda, comenzando en A1) y responde en la columna B de cada fila. Dedica entre 2 y 3 minutos a esta sección.

   **Pregunta 1 (A1):**
   ```
   ¿Cuál modelo (PMS-3 o PMS-5) reacciona más rápido ante un aumento súbito de demanda? ¿Por qué?
   ```
   
   **Pregunta 2 (A3):**
   ```
   Si el producto que analizaste tiene una temporada alta en las semanas 18-22 (mayor demanda), ¿cuál modelo anticiparía mejor ese pico? Justifica con base en los gráficos.
   ```
   
   **Pregunta 3 (A5):**
   ```
   En términos de inventario de seguridad: ¿un modelo con mayor error de pronóstico requiere más o menos stock de seguridad? ¿Por qué?
   ```
   
   **Pregunta 4 (A7):**
   ```
   Si eres responsable de planificar el picking semanal en una bodega, ¿qué consecuencia operativa tendría usar PMS-5 en lugar de PMS-3 cuando la demanda cambia repentinamente?
   ```
   
   **Pregunta 5 (A9):**
   ```
   ¿En qué tipo de producto logístico (alta rotación estable, producto estacional, producto de demanda errática) recomendarías usar PMS-3 vs PMS-5? Da un ejemplo concreto de cada caso.
   ```

3. Escribe tus respuestas en las celdas **B1, B3, B5, B7 y B9** respectivamente. Las respuestas pueden ser breves (2-4 oraciones cada una).

4. Guarda el archivo final.

#### Resultado esperado

La hoja "Reflexión" debe contener 5 preguntas con sus respectivas respuestas escritas por el participante. No existe una única respuesta correcta, pero las respuestas deben demostrar comprensión del trade-off entre reactividad y estabilidad del pronóstico, y su impacto en decisiones de inventario y picking.

#### Guía de respuestas esperadas (referencia para el instructor)

| Pregunta | Respuesta esperada (síntesis) |
|---|---|
| 1 | PMS-3 reacciona más rápido porque su ventana es más corta; los datos recientes tienen mayor peso relativo en el promedio. |
| 2 | PMS-3, porque al usar solo 3 semanas, comenzará a capturar el aumento de demanda antes que PMS-5, que sigue "arrastrando" semanas de baja demanda en su ventana. |
| 3 | Mayor error requiere mayor stock de seguridad, porque la incertidumbre sobre la demanda real es mayor y el riesgo de quiebre de stock aumenta. |
| 4 | Con PMS-5, el equipo de picking podría preparar menos unidades de las necesarias si la demanda sube abruptamente, generando retrasos en los despachos y potencial quiebre de stock. |
| 5 | PMS-5 para productos de alta rotación estable (ej. agua embotellada). PMS-3 para productos con variabilidad moderada o tendencia cambiante (ej. productos de temporada escolar). |

#### Verificación

- [ ] La hoja "Reflexión" existe y está correctamente nombrada.
- [ ] Las 5 preguntas están escritas en la columna A.
- [ ] Se han escrito respuestas (aunque sean breves) en la columna B para cada pregunta.
- [ ] El archivo final ha sido guardado.

---

## Validación y Pruebas

Antes de entregar tu archivo, realiza las siguientes verificaciones finales de integridad:

### Lista de verificación final

| Elemento | Criterio de validación | ✅/❌ |
|---|---|---|
| Estructura de la tabla | Columnas A-F con encabezados en negrita, datos en filas 2-25 | |
| PMS-3 completo | Celdas C5:C25 con valores numéricos; C2:C4 vacías | |
| PMS-5 completo | Celdas D7:D25 con valores numéricos; D2:D6 vacías | |
| Fórmulas correctas | C5=`PROMEDIO(B2:B4)`, D7=`PROMEDIO(B2:B6)`, desplazamiento correcto | |
| Errores calculados | E5:E25 y F7:F25 con valores; sin errores de fórmula | |
| Gráfico de líneas | 3 series, eje X con S1-S24, leyenda visible, título correcto | |
| Tabla de indicadores | 4 indicadores calculados para PMS-3 y PMS-5 | |
| Hoja Reflexión | 5 preguntas con respuestas escritas | |
| Nombre del archivo | `Lab13_PMS_[TuNombre].xlsx` | |

### Prueba de consistencia numérica

Para verificar que tus fórmulas son correctas, realiza este cálculo manual y compáralo con Excel:

1. Toma los valores reales de las semanas S1, S2 y S3 de tu dataset.
2. Calcula manualmente: `(S1 + S2 + S3) / 3`
3. Este resultado debe coincidir exactamente con el valor en la celda **C5** de tu tabla.
4. Repite con S1, S2, S3, S4, S5 y verifica que coincida con **D7**.

Si los valores no coinciden, revisa que las referencias de celda en tus fórmulas apunten a las filas correctas.

---

## Solución de Problemas

### Problema 1: Las fórmulas de PMS se desplazan incorrectamente al arrastrar

**Síntoma:** Al arrastrar la fórmula de C5 hacia abajo, las celdas intermedias muestran valores inesperadamente altos o bajos, o aparecen errores `#¡VALOR!`. Al hacer clic en C10, por ejemplo, la fórmula muestra `=PROMEDIO(B7:B9)` en lugar de `=PROMEDIO(B7:B9)` (correcto) o muestra `=PROMEDIO($B$2:$B$4)` (incorrecto — rango fijo).

**Causa:** El uso accidental de referencias absolutas (`$B$2:$B$4`) en lugar de referencias relativas (`B2:B4`) impide que el rango se desplace al arrastrar. Esto ocurre si se presionó `F4` durante la escritura de la fórmula, o si se copió desde otra fuente que usaba referencias absolutas.

**Solución:**
1. Haz clic en la celda **C5**.
2. Revisa la barra de fórmulas: debe mostrar `=PROMEDIO(B2:B4)` **sin** signos `$`.
3. Si hay signos `$`, edita la fórmula eliminándolos manualmente: `=PROMEDIO(B2:B4)`.
4. Presiona `Enter` y vuelve a arrastrar desde C5 hasta C25.
5. Verifica C10: debe mostrar `=PROMEDIO(B7:B9)` (el rango se desplazó 5 filas hacia abajo, igual que la celda).

---

### Problema 2: El gráfico muestra las series PMS-3 y PMS-5 comenzando desde el mismo punto que la demanda real (no respeta las celdas vacías)

**Síntoma:** En el gráfico, las tres líneas comienzan desde S1 (semana 1), y las series PMS-3 y PMS-5 muestran valores de 0 (cero) o valores incorrectos en las primeras semanas donde deberían estar vacías.

**Causa:** Excel interpreta las celdas vacías de C2:C4 y D2:D6 como cero cuando construye el gráfico, especialmente si en algún momento se escribió y luego se borró un valor en esas celdas (dejando formato residual). También puede ocurrir si el rango del gráfico fue definido incorrectamente incluyendo filas con ceros.

**Solución:**
1. Selecciona el rango **C2:C4** y presiona `Supr` (Delete) para asegurarte de que estén completamente vacías (no solo con contenido borrado visualmente).
2. Repite con **D2:D6**.
3. Haz clic derecho sobre el gráfico → **Seleccionar datos**.
4. En el panel izquierdo, selecciona la serie **PMS-3** y haz clic en **Editar**.
5. Verifica que el rango de valores sea `=Datos!$C$2:$C$25`. Haz clic en Aceptar.
6. Para controlar cómo Excel muestra las celdas vacías: haz clic en **Celdas ocultas y vacías** (en el panel Seleccionar datos) → selecciona **Espacios** (en lugar de "Cero"). Esto hará que las líneas comiencen correctamente donde hay datos.
7. Haz clic en **Aceptar** y verifica que las líneas ahora comiencen en S4 (PMS-3) y S6 (PMS-5).

---

## Limpieza y Entrega

### Pasos de limpieza antes de entregar

1. **Elimina cualquier cálculo provisional** que hayas realizado en celdas fuera de la estructura definida (columnas G en adelante o filas por debajo de la tabla de indicadores).

2. **Verifica que no haya hojas innecesarias:** El archivo final debe contener exactamente dos hojas: **"Datos"** y **"Reflexión"**. Si creaste hojas adicionales de prueba, elimínalas haciendo clic derecho sobre la pestaña → **Eliminar**.

3. **Ajusta el ancho de columnas** para que todos los valores sean legibles sin truncamiento:
   - Selecciona las columnas A a F.
   - Haz doble clic en cualquier borde entre encabezados de columna para autoajustar el ancho.

4. **Confirma el nombre del archivo:** Debe ser `Lab13_PMS_[TuNombre].xlsx`. Reemplaza `[TuNombre]` con tu nombre real (ej. `Lab13_PMS_MariaGomez.xlsx`).

5. **Guarda el archivo por última vez:** `Ctrl + G` (Windows) o `Cmd + S` (Mac).

### Entrega

Sube el archivo `Lab13_PMS_[TuNombre].xlsx` a la plataforma indicada por el instructor. El archivo debe contener:
- ✅ Hoja "Datos" con tabla completa (columnas A–F) y gráfico de líneas.
- ✅ Tabla de indicadores estadísticos (filas 28–32).
- ✅ Hoja "Reflexión" con las 5 preguntas respondidas.

---

## Resumen

En esta práctica construiste manualmente dos modelos de Promedio Móvil Simple en Excel — PMS-3 y PMS-5 — aplicando la fórmula `PROMEDIO()` sobre rangos desplazados para generar pronósticos semana a semana. Los puntos clave que consolidaste son:

| Concepto | Aprendizaje clave |
|---|---|
| **Construcción del PMS** | Los primeros N−1 períodos no tienen pronóstico; la fórmula se desplaza una fila por período usando referencias relativas. |
| **Efecto del tamaño de ventana** | N pequeño (PMS-3) → más reactivo, mayor variabilidad en errores. N grande (PMS-5) → más suave, mayor rezago ante cambios. |
| **Error simple** | Demanda Real − Pronóstico. Errores positivos = subestimación; negativos = sobreestimación. |
| **Implicación logística** | El tamaño de N debe elegirse según el perfil de demanda del producto: PMS-3 para demanda volátil, PMS-5 para demanda estable. |
| **Stock de seguridad** | Mayor variabilidad en los errores de pronóstico implica mayor necesidad de stock de seguridad para mantener el nivel de servicio. |

### Conexión con las próximas prácticas

Los conceptos de ventana de suavización y trade-off reactividad/estabilidad que trabajaste aquí son la base para entender el **Promedio Móvil Ponderado** (Lección 3.2), donde los períodos más recientes reciben mayor peso, y el **Suavizamiento Exponencial** (Lección 3.3), que lleva esta lógica a su forma más eficiente. Los errores calculados en esta práctica también servirán como punto de comparación cuando calcules MAE y MAPE en prácticas posteriores del módulo.

### Recursos de referencia

| Recurso | Descripción |
|---|---|
| [Hyndman & Athanasopoulos — *Forecasting: Principles and Practice*](https://otexts.com/fpp3/) | Capítulo 3: Moving Average Methods (acceso gratuito en línea) |
| [Microsoft Support — Función PROMEDIO](https://support.microsoft.com/es-es/office/promedio-funci%C3%B3n-promedio-047bac88-d466-426c-a32b-8f33eb960cf6) | Documentación oficial de la función PROMEDIO en Excel |
| [Microsoft Support — Crear un gráfico de líneas](https://support.microsoft.com/es-es/office/crear-un-gr%C3%A1fico-de-l%C3%ADneas-3d6a0d31-4c2f-4c55-b7c6-09c7a43e1ee0) | Guía paso a paso para gráficos de líneas en Excel |
| Chopra & Meindl — *Supply Chain Management* (Pearson) | Capítulo 7: Demand Forecasting in a Supply Chain |

---

---

---LAB_START---
LAB_ID: 03-00-02
---MARKDOWN---
# Práctica 14 — Aplicación de Promedio Móvil Ponderado

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 3 — Métodos de Pronóstico Cuantitativo |
| **Práctica anterior** | Práctica 13 — Promedio Móvil Simple (PMS-3) |
| **Herramienta principal** | Microsoft Excel 2016 o superior |

---

## 2. Descripción General

En la Práctica 13 construiste el modelo de Promedio Móvil Simple con ventana de 3 períodos (PMS-3), donde cada uno de los tres períodos anteriores aportaba exactamente el mismo peso al pronóstico. En esta práctica darás un paso adelante: implementarás el **Promedio Móvil Ponderado (PMP)**, un método que asigna mayor importancia a los períodos más recientes, reconociendo que la demanda reciente contiene más información relevante sobre el comportamiento futuro que la demanda más antigua.

Trabajarás con dos esquemas de pesos (Esquema A y Esquema B), calcularás el MAE preliminar para los tres modelos (PMS-3, PMP-A y PMP-B), y realizarás un ejercicio de sensibilidad para desarrollar intuición sobre cómo la elección de pesos impacta la precisión del pronóstico. Al finalizar, podrás justificar operativamente cuál método conviene usar para el producto analizado en el contexto de decisiones de reabastecimiento e inventario.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica, serás capaz de:

- [ ] **Implementar** el Promedio Móvil Ponderado (PMP) en Excel usando `SUMAPRODUCTO`, asignando pesos diferenciados a cada período de la ventana de 3 semanas.
- [ ] **Verificar** que la suma de pesos sea igual a 1 mediante una celda de validación, garantizando la consistencia del modelo.
- [ ] **Calcular** el MAE preliminar para los modelos PMS-3, PMP-A y PMP-B, y construir una tabla comparativa de desempeño.
- [ ] **Analizar** el efecto de modificar los pesos sobre el MAE en un ejercicio de sensibilidad, desarrollando intuición sobre la optimización de parámetros.
- [ ] **Seleccionar** y justificar operativamente el modelo con menor error para apoyar decisiones de reabastecimiento e inventario.

---

## 4. Prerrequisitos

### Conocimientos previos
- Haber completado la **Práctica 13** y contar con el archivo Excel resultante que contiene la serie de 24 semanas y el modelo PMS-3 construido con sus errores absolutos calculados.
- Comprensión del concepto de **promedio ponderado** (contenido teórico del Módulo 3.2): saber que la suma de todos los pesos debe ser igual a 1 y que el peso mayor se asigna al período más reciente.
- Conocimiento básico de la función `SUMAPRODUCTO` en Excel: entender que multiplica elemento a elemento dos rangos y suma los resultados.
- Familiaridad con la función `ABS` para calcular errores absolutos, trabajada en la Práctica 13.

### Acceso y archivos necesarios
- Archivo Excel de la Práctica 13 (nombre sugerido: `Practica13_PMS_TuNombre.xlsx`) con las columnas: Semana, Demanda Real, PMS-3, Error Absoluto PMS-3.
- Microsoft Excel 2016 o superior instalado y funcional.
- En caso de no contar con el archivo de la Práctica 13, solicitar al instructor el **archivo checkpoint** `Checkpoint_P13_Completo.xlsx`.

---

## 5. Entorno de Laboratorio

### Hardware requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| Memoria RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Espacio en disco | 2 GB libres | 2 GB libres |
| Conexión a internet | Opcional para esta práctica | — |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado |
| Archivo de datos | `Practica13_PMS_TuNombre.xlsx` | Generado en Práctica 13 |

### Preparación del entorno

Antes de comenzar, realiza las siguientes acciones:

1. Abre el archivo `Practica13_PMS_TuNombre.xlsx` desde tu carpeta de trabajo.
2. Guarda inmediatamente una copia con el nuevo nombre:

```
Practica14_PMP_TuNombre.xlsx
```

3. Verifica que la hoja contenga al menos las siguientes columnas en la **Hoja1** (o la hoja de trabajo de la Práctica 13):
   - Columna A: `Semana` (S1 a S24)
   - Columna B: `Demanda Real` (24 valores numéricos)
   - Columna C: `PMS-3` (pronósticos desde S4 en adelante)
   - Columna D: `Error Absoluto PMS-3`

4. Asegúrate de que las filas de datos comiencen en la **fila 2** (fila 1 = encabezados). Esta práctica asume esa estructura. Si tu archivo tiene una estructura diferente, ajusta las referencias de celda en cada paso.

> **Nota de referencia rápida:** La función `SUMAPRODUCTO` tiene la siguiente sintaxis:
> `=SUMAPRODUCTO(array1, array2)` — multiplica cada par de elementos y suma todos los productos.
> Ejemplo: `=SUMAPRODUCTO({0.5,0.3,0.2},{100,110,90})` = 0.5×100 + 0.3×110 + 0.2×90 = **101**

---

## 6. Instrucciones Paso a Paso

---

### Paso 1 — Preparar la estructura de columnas para PMP

**Objetivo:** Crear los encabezados y la zona de validación de pesos antes de ingresar fórmulas, asegurando una hoja organizada y auditable.

#### Instrucciones

1. Abre el archivo `Practica14_PMP_TuNombre.xlsx` que guardaste en el paso de preparación.

2. En la **fila 1**, agrega los siguientes encabezados en las columnas indicadas (las columnas A–D ya existen de la Práctica 13):

   | Celda | Encabezado |
   |---|---|
   | E1 | `PMP-A (0.5/0.3/0.2)` |
   | F1 | `Error Abs PMP-A` |
   | G1 | `PMP-B (0.6/0.3/0.1)` |
   | H1 | `Error Abs PMP-B` |

3. Ahora crea la **zona de parámetros y validación de pesos**. Utiliza un área separada de la tabla principal, por ejemplo a partir de la celda **J1**:

   | Celda | Contenido |
   |---|---|
   | J1 | `PARÁMETROS DE PESOS` |
   | J2 | `Esquema A - Peso t (más reciente)` |
   | K2 | `0.5` |
   | J3 | `Esquema A - Peso t-1` |
   | K3 | `0.3` |
   | J4 | `Esquema A - Peso t-2 (más antiguo)` |
   | K4 | `0.2` |
   | J5 | `SUMA Esquema A (debe = 1)` |
   | K5 | `=K2+K3+K4` |
   | J7 | `Esquema B - Peso t (más reciente)` |
   | K7 | `0.6` |
   | J8 | `Esquema B - Peso t-1` |
   | K8 | `0.3` |
   | J9 | `Esquema B - Peso t-2 (más antiguo)` |
   | K9 | `0.1` |
   | J10 | `SUMA Esquema B (debe = 1)` |
   | K10 | `=K7+K8+K9` |

4. Aplica formato condicional a las celdas K5 y K10 para que se resalten en **rojo** si el valor es diferente de 1:
   - Selecciona K5 → **Inicio → Formato condicional → Nueva regla → Usar fórmula**
   - Ingresa la fórmula: `=K5<>1`
   - Elige relleno rojo y texto blanco → Aceptar.
   - Repite el mismo proceso para K10.

5. Verifica que K5 muestre `1` y K10 muestre `1`. Si alguna celda muestra un valor diferente, corrige los pesos antes de continuar.

#### Resultado esperado
La hoja debe mostrar los encabezados de columnas E–H en la fila 1 y la zona de parámetros en la columna J–K con las sumas de pesos en verde (o sin resaltar en rojo), confirmando que ambos esquemas suman exactamente 1.

#### Verificación
Comprueba visualmente que:
- [ ] K5 = `1` (sin relleno rojo)
- [ ] K10 = `1` (sin relleno rojo)
- [ ] Los encabezados E1:H1 están correctamente escritos

---

### Paso 2 — Implementar PMP Esquema A con SUMAPRODUCTO

**Objetivo:** Construir la fórmula del Promedio Móvil Ponderado con pesos (0.5, 0.3, 0.2) usando `SUMAPRODUCTO`, y extenderla para las 22 semanas pronosticables.

#### Instrucciones

1. Haz clic en la celda **E4** (primera fila donde es posible calcular un PMP con ventana de 3 períodos, ya que necesita S1, S2 y S3 como historia).

2. Ingresa la siguiente fórmula. Esta fórmula multiplica cada valor de demanda por su peso correspondiente: el más reciente (B4-1 = B3... espera, lee el orden cuidadosamente):

   > **Orden de pesos importante:** El peso **mayor (0.5) corresponde al período más reciente (t)**. Para el pronóstico de S4, los datos son: t = S3 (B4), t-1 = S2 (B3), t-2 = S1 (B2). Sin embargo, como el pronóstico de S4 usa los datos de S1, S2 y S3, la celda E4 usa B2:B4 con pesos en orden **antiguo→reciente = 0.2, 0.3, 0.5**.

   ```excel
   =SUMAPRODUCTO(B2:B4, $K$4:$K$2)
   ```

   > **Atención:** La referencia `$K$4:$K$2` es un rango inverso que Excel **no acepta directamente**. Usa en su lugar la forma explícita con los valores de la zona de parámetros en el orden correcto. La fórmula correcta y recomendada es:

   ```excel
   =SUMAPRODUCTO({1,1,1}*B2:B4, {0.2,0.3,0.5})
   ```

   O mejor aún, usando referencias absolutas a los parámetros para que el ejercicio de sensibilidad del Paso 5 funcione automáticamente:

   ```excel
   =SUMAPRODUCTO(B2:B4, {0.2,0.3,0.5})
   ```

   > **Forma más robusta (recomendada):** Usa referencias a las celdas de parámetros en orden antiguo→reciente:

   ```excel
   =$K$4*B2 + $K$3*B3 + $K$2*B4
   ```

   Esta forma explícita es la más legible y permite que el ejercicio de sensibilidad del Paso 5 funcione correctamente al modificar K2, K3 y K4.

3. Presiona **Enter** y verifica que el resultado sea un número razonable (cercano al promedio de B2:B4 pero desplazado hacia B4 por los pesos).

4. Haz clic nuevamente en **E4** y arrastra la fórmula hacia abajo hasta la celda **E25** (pronóstico para S25, usando datos de S22, S23 y S24). Esto generará pronósticos para S4 hasta S25.

   > **Nota:** S25 es el pronóstico para la semana siguiente al final del historial (semana 25). Es el pronóstico operativo más relevante.

5. Deja las celdas **E2 y E3** vacías (no hay suficiente historia para calcular PMP con ventana de 3 en esas filas). Opcionalmente, escribe `—` o deja en blanco.

6. Aplica formato numérico a E4:E25 con **1 decimal**: selecciona el rango → `Ctrl+1` → Número → Decimales: 1.

#### Resultado esperado
La columna E debe mostrar valores numéricos desde E4 hasta E25, representando los pronósticos PMP-A. Los valores deben ser similares a los de PMS-3 (columna C) pero con ligeras diferencias, reflejando el mayor peso en el período más reciente.

#### Verificación
Calcula manualmente el valor de E4 para validar la fórmula:
- Datos: B2 (S1), B3 (S2), B4 (S3) con pesos 0.2, 0.3, 0.5 respectivamente.
- Fórmula manual: `0.2 × [valor S1] + 0.3 × [valor S2] + 0.5 × [valor S3]`
- Compara con el valor en E4. Deben coincidir exactamente.
- [ ] E4 coincide con el cálculo manual
- [ ] E25 contiene un valor (pronóstico para semana 25)

---

### Paso 3 — Implementar PMP Esquema B con SUMAPRODUCTO

**Objetivo:** Replicar el proceso del Paso 2 para el Esquema B con pesos (0.6, 0.3, 0.1), aplicando mayor énfasis en el período más reciente.

#### Instrucciones

1. Haz clic en la celda **G4**.

2. Ingresa la fórmula usando las referencias a los parámetros del Esquema B (K7 = 0.6 para el más reciente, K8 = 0.3, K9 = 0.1 para el más antiguo):

   ```excel
   =$K$9*B2 + $K$8*B3 + $K$7*B4
   ```

   > El orden es: K9 (peso del más antiguo, t-2) × B2 + K8 (peso de t-1) × B3 + K7 (peso del más reciente, t) × B4.

3. Presiona **Enter** y verifica el resultado.

4. Arrastra la fórmula desde **G4** hasta **G25**.

5. Deja G2 y G3 vacías (igual que en la columna E).

6. Aplica el mismo formato numérico de 1 decimal al rango G4:G25.

7. **Verificación cruzada rápida:** Para cualquier fila del rango, el PMP-B debería estar más "pegado" al valor más reciente que el PMP-A, porque el peso de 0.6 en el período t es mayor que el 0.5 del Esquema A. Observa si esto se cumple visualmente en algunas filas.

#### Resultado esperado
La columna G contiene los pronósticos PMP-B desde G4 hasta G25. Los valores en G deben diferir de los de E, siendo el Esquema B generalmente más reactivo a cambios recientes en la demanda.

#### Verificación
Calcula manualmente el valor de G4:
- Fórmula: `0.1 × [valor S1] + 0.3 × [valor S2] + 0.6 × [valor S3]`
- [ ] G4 coincide con el cálculo manual
- [ ] G4 ≠ E4 (los esquemas deben producir valores diferentes)
- [ ] G25 contiene un valor numérico

---

### Paso 4 — Calcular Errores Absolutos y MAE para los Tres Modelos

**Objetivo:** Calcular el error absoluto período a período para PMP-A y PMP-B, y obtener el MAE preliminar para los tres modelos (PMS-3, PMP-A y PMP-B) para compararlos.

#### Instrucciones

**Parte A: Errores absolutos de PMP-A**

1. Haz clic en la celda **F4**.

2. Ingresa la fórmula del error absoluto comparando el pronóstico PMP-A contra la demanda real del mismo período:

   ```excel
   =ABS(B4-E4)
   ```

   > El error se calcula comparando el pronóstico del período t (calculado con datos hasta t-1) contra la demanda real del período t.

3. Arrastra la fórmula desde **F4** hasta **F24** (no hasta F25, porque S25 es el pronóstico futuro y no tiene demanda real aún).

4. Deja F2, F3 y F25 vacías.

**Parte B: Errores absolutos de PMP-B**

5. Haz clic en la celda **H4**.

6. Ingresa la misma lógica para el Esquema B:

   ```excel
   =ABS(B4-G4)
   ```

7. Arrastra desde **H4** hasta **H24**.

**Parte C: Tabla resumen de MAE**

8. Crea una tabla resumen en el área debajo de los parámetros, por ejemplo a partir de la celda **J13**:

   | Celda | Contenido |
   |---|---|
   | J13 | `TABLA COMPARATIVA DE MODELOS` |
   | J14 | `Modelo` |
   | K14 | `MAE Preliminar` |
   | J15 | `PMS-3` |
   | K15 | `=PROMEDIO(D4:D24)` |
   | J16 | `PMP-A (0.5/0.3/0.2)` |
   | K16 | `=PROMEDIO(F4:F24)` |
   | J17 | `PMP-B (0.6/0.3/0.1)` |
   | K17 | `=PROMEDIO(H4:H24)` |
   | J18 | `Mejor modelo (menor MAE)` |
   | K18 | `=INDICE(J15:J17,COINCIDIR(MIN(K15:K17),K15:K17,0))` |

   > La fórmula en K18 identifica automáticamente el nombre del modelo con menor MAE usando `INDICE` y `COINCIDIR`.

9. Aplica formato numérico con **2 decimales** a las celdas K15:K17.

10. Aplica **negrita y relleno de color amarillo claro** a la fila del modelo con menor MAE (puedes hacerlo manualmente después de ver los resultados, o usar formato condicional).

    Para formato condicional automático en K15:K17:
    - Selecciona K15:K17
    - **Inicio → Formato condicional → Nueva regla → Usar fórmula**
    - Fórmula: `=K15=MIN($K$15:$K$17)`
    - Relleno: amarillo → Aceptar
    - Repite para K16 con `=K16=MIN($K$15:$K$17)` y K17 con `=K17=MIN($K$15:$K$17)`

    > Alternativa más simple: selecciona K15:K17 → **Inicio → Formato condicional → Resaltar reglas de celdas → Igual a** → `=MIN($K$15:$K$17)`.

#### Resultado esperado
- Las columnas F4:F24 y H4:H24 contienen errores absolutos para cada período.
- La tabla en J13:K18 muestra los tres MAE calculados y el nombre del mejor modelo resaltado.
- Los valores de MAE deben ser números positivos. Si algún MAE es 0, revisa que las fórmulas de pronóstico no estén referenciando directamente la demanda real.

#### Verificación
- [ ] F4:F24 contiene 21 valores numéricos (errores absolutos PMP-A)
- [ ] H4:H24 contiene 21 valores numéricos (errores absolutos PMP-B)
- [ ] K15, K16 y K17 muestran valores numéricos positivos
- [ ] K18 muestra el nombre de uno de los tres modelos
- [ ] Los tres MAE son distintos entre sí (si son idénticos, hay un error en las fórmulas)

---

### Paso 5 — Ejercicio de Sensibilidad: Modificación de Pesos

**Objetivo:** Desarrollar intuición sobre la optimización de parámetros observando cómo el MAE cambia al modificar los pesos del Esquema A manualmente, sin alterar la estructura de fórmulas.

#### Instrucciones

> **Importante:** Este ejercicio modifica temporalmente los pesos del Esquema A en la zona de parámetros. Gracias a que las fórmulas de la columna E referencian las celdas K2, K3 y K4, el MAE se recalculará automáticamente al cambiar los pesos. Al finalizar el ejercicio, **restaurarás los pesos originales**.

**Preparación: Registrar valores originales**

1. Antes de modificar nada, anota (en papel o en una celda auxiliar) los valores actuales:
   - K15 (MAE PMS-3): `___________`
   - K16 (MAE PMP-A original): `___________`
   - K17 (MAE PMP-B): `___________`

2. Crea una pequeña tabla de registro de sensibilidad en la celda **J21**:

   | Celda | Contenido |
   |---|---|
   | J21 | `REGISTRO DE SENSIBILIDAD - ESQUEMA A` |
   | J22 | `Pesos (t / t-1 / t-2)` |
   | K22 | `MAE PMP-A` |
   | J23 | `0.5 / 0.3 / 0.2 (original)` |
   | K23 | `=K16` (referencia dinámica) |

**Prueba 1: Pesos más concentrados en el período reciente**

3. Modifica los pesos del Esquema A a los siguientes valores:
   - K2 (peso t, más reciente): `0.7`
   - K3 (peso t-1): `0.2`
   - K4 (peso t-2, más antiguo): `0.1`

4. Verifica que K5 siga mostrando `1` (la suma de pesos debe mantenerse igual a 1). Si no es 1, ajusta los valores.

5. Observa el nuevo valor de K16 (MAE PMP-A). Anótalo en la tabla de registro:
   - En J24 escribe: `0.7 / 0.2 / 0.1`
   - En K24 escribe el valor actual de K16 (escríbelo como número fijo, no como fórmula, para preservar el registro).

**Prueba 2: Pesos más uniformes (casi como PMS)**

6. Modifica los pesos del Esquema A a:
   - K2 (peso t): `0.4`
   - K3 (peso t-1): `0.35`
   - K4 (peso t-2): `0.25`

7. Verifica K5 = 1. Anota el nuevo MAE:
   - En J25: `0.4 / 0.35 / 0.25`
   - En K25: valor actual de K16.

**Prueba 3: Pesos invertidos (mayor peso al más antiguo — no recomendado, pero instructivo)**

8. Modifica los pesos del Esquema A a:
   - K2 (peso t): `0.2`
   - K3 (peso t-1): `0.3`
   - K4 (peso t-2): `0.5`

9. Verifica K5 = 1. Anota el MAE:
   - En J26: `0.2 / 0.3 / 0.5 (invertido)`
   - En K26: valor actual de K16.

**Restauración de valores originales**

10. Restaura los pesos originales del Esquema A:
    - K2: `0.5`
    - K3: `0.3`
    - K4: `0.2`

11. Verifica que K5 = 1 y que K16 muestre el mismo MAE que anotaste en el paso 1.

**Análisis reflexivo**

12. Observa tu tabla de registro (J21:K26) y responde mentalmente (o escribe en una celda de comentario):
    - ¿Qué combinación de pesos produjo el menor MAE?
    - ¿Los pesos más concentrados en el período reciente siempre mejoran el pronóstico?
    - ¿Qué sucedió con el MAE cuando los pesos estaban invertidos?

#### Resultado esperado
La tabla de registro muestra 4 combinaciones de pesos con sus MAE correspondientes. Los valores de MAE varían entre combinaciones, mostrando que la elección de pesos tiene impacto real en la precisión. En la mayoría de series de demanda con cierta variabilidad, los pesos invertidos producirán el MAE más alto.

#### Verificación
- [ ] K5 = `1` después de restaurar los pesos originales
- [ ] K16 volvió al valor original después de la restauración
- [ ] La tabla de registro en J21:K26 tiene 4 filas de datos completas
- [ ] Los 4 valores de MAE en el registro son diferentes entre sí

---

### Paso 6 — Crear Gráfico Comparativo de los Tres Modelos

**Objetivo:** Visualizar en un gráfico de líneas la demanda real frente a los tres pronósticos, para interpretar visualmente las diferencias de comportamiento entre PMS-3, PMP-A y PMP-B.

#### Instrucciones

1. Selecciona el rango **A1:C24** (Semana, Demanda Real y PMS-3). Mantén presionado **Ctrl** y selecciona adicionalmente **E4:E24** y **G4:G24**.

   > **Alternativa más simple para el gráfico:** Primero crea el gráfico con solo Demanda Real y PMS-3, y luego agrega las series PMP-A y PMP-B.

2. Ve a **Insertar → Gráficos → Gráfico de líneas → Línea con marcadores**.

3. Excel generará un gráfico. Haz clic derecho sobre el gráfico → **Seleccionar datos**.

4. Verifica que aparezcan las siguientes series (agrégalas si no están):
   - Serie 1: `Demanda Real` (columna B, filas 2–24)
   - Serie 2: `PMS-3` (columna C, filas 4–24)
   - Serie 3: `PMP-A` (columna E, filas 4–24)
   - Serie 4: `PMP-B` (columna G, filas 4–24)

5. Configura el eje horizontal con las etiquetas de semana (columna A).

6. Personaliza el gráfico:
   - **Título del gráfico:** `Comparación de Modelos de Pronóstico: PMS-3 vs PMP-A vs PMP-B`
   - **Demanda Real:** línea negra gruesa (grosor 2.5 pt)
   - **PMS-3:** línea azul punteada
   - **PMP-A:** línea naranja
   - **PMP-B:** línea verde
   - Activa la **leyenda** en la parte inferior del gráfico.

7. Mueve el gráfico a una posición que no tape la tabla de datos. Sugerencia: colócalo a la derecha de la columna K o en una hoja nueva llamada `Gráfico_Comparativo`.

8. Observa el gráfico y responde:
   - ¿Cuál modelo sigue más de cerca la demanda real?
   - ¿Se aprecia visualmente el rezago del PMS-3 frente a los modelos ponderados?
   - ¿En qué semanas la diferencia entre modelos es más notable?

#### Resultado esperado
Un gráfico de líneas con 4 series claramente diferenciadas por color, mostrando las 24 semanas de historial. Los modelos PMP deben verse ligeramente más "pegados" a la demanda real que el PMS-3 en los períodos de cambio de tendencia.

#### Verificación
- [ ] El gráfico tiene título
- [ ] Aparecen las 4 series (Demanda Real, PMS-3, PMP-A, PMP-B)
- [ ] La leyenda es visible y correcta
- [ ] El eje horizontal muestra las semanas (S1–S24 o equivalente)

---

### Paso 7 — Interpretación Operativa y Documentación Final

**Objetivo:** Traducir los resultados estadísticos en una recomendación operativa concreta para decisiones de reabastecimiento, completando el ciclo de análisis de la práctica.

#### Instrucciones

1. Crea una nueva hoja en el archivo llamada `Recomendacion_Operativa`.

2. En esa hoja, construye la siguiente tabla de conclusiones (puedes copiarla y completar los campos en blanco con tus resultados reales):

   ```
   ╔══════════════════════════════════════════════════════════════════╗
   ║         REPORTE DE SELECCIÓN DE MODELO DE PRONÓSTICO            ║
   ║                     Práctica 14 — PMP                           ║
   ╠══════════════════════════════════════════════════════════════════╣
   ║ Producto analizado:    [Nombre/SKU del dataset]                  ║
   ║ Período de análisis:   Semanas 1 a 24                           ║
   ║ Fecha de análisis:     [Fecha actual]                           ║
   ╠══════════════════════════════════════════════════════════════════╣
   ║ RESULTADOS DE MAE                                               ║
   ║  PMS-3:          [valor]                                        ║
   ║  PMP-A (0.5/0.3/0.2): [valor]                                  ║
   ║  PMP-B (0.6/0.3/0.1): [valor]                                  ║
   ║  MEJOR MODELO:   [nombre del modelo con menor MAE]              ║
   ╠══════════════════════════════════════════════════════════════════╣
   ║ PRONÓSTICO SEMANA 25 (modelo seleccionado):  [valor] unidades   ║
   ╠══════════════════════════════════════════════════════════════════╣
   ║ IMPLICACIÓN OPERATIVA:                                          ║
   ║  Con base en el pronóstico de [X] unidades para la semana 25,   ║
   ║  se recomienda:                                                 ║
   ║  - Orden de reabastecimiento: [X + buffer de seguridad] u.      ║
   ║  - Preparación de picking: estimar [X] unidades para            ║
   ║    operaciones de bodega en semana 25.                          ║
   ║  - Nivel de servicio: si el CV del producto es < 0.3,           ║
   ║    el modelo PMP es suficiente; si CV > 0.3, considerar         ║
   ║    suavizamiento exponencial (Lección 3.3).                     ║
   ╠══════════════════════════════════════════════════════════════════╣
   ║ JUSTIFICACIÓN DE SELECCIÓN DEL MODELO:                          ║
   ║  [Escribe 2-3 oraciones explicando por qué el modelo con menor  ║
   ║   MAE es el más adecuado para este producto, considerando el    ║
   ║   comportamiento observado en el gráfico comparativo]           ║
   ╚══════════════════════════════════════════════════════════════════╝
   ```

3. En la celda correspondiente al **pronóstico semana 25**, referencia directamente la celda del modelo ganador:
   - Si ganó PMP-A: `=Hoja1!E25`
   - Si ganó PMP-B: `=Hoja1!G25`
   - Si ganó PMS-3: `=Hoja1!C25`

4. Guarda el archivo: `Ctrl + S`.

#### Resultado esperado
Una hoja de recomendación operativa con todos los campos completados, que conecta el análisis estadístico (MAE) con una decisión concreta de reabastecimiento y planificación de picking para la semana 25.

#### Verificación
- [ ] La hoja `Recomendacion_Operativa` existe en el archivo
- [ ] Todos los campos de la tabla están completados con valores reales (no con corchetes vacíos)
- [ ] El pronóstico de la semana 25 está referenciado dinámicamente desde la hoja de datos
- [ ] El archivo está guardado con el nombre correcto

---

## 7. Validación y Pruebas

Una vez completados todos los pasos, realiza las siguientes verificaciones finales para confirmar que el laboratorio está correctamente ejecutado:

### Lista de verificación final

| # | Verificación | Criterio de éxito |
|---|---|---|
| 1 | **Suma de pesos Esquema A** | K5 = 1.0 exacto |
| 2 | **Suma de pesos Esquema B** | K10 = 1.0 exacto |
| 3 | **Pronósticos PMP-A** | E4:E25 contiene 22 valores numéricos |
| 4 | **Pronósticos PMP-B** | G4:G25 contiene 22 valores numéricos |
| 5 | **Errores PMP-A** | F4:F24 contiene 21 valores ≥ 0 |
| 6 | **Errores PMP-B** | H4:H24 contiene 21 valores ≥ 0 |
| 7 | **MAE PMS-3** | K15 es un número positivo |
| 8 | **MAE PMP-A** | K16 es un número positivo y ≠ K15 |
| 9 | **MAE PMP-B** | K17 es un número positivo y ≠ K15 ≠ K16 |
| 10 | **Mejor modelo identificado** | K18 muestra el nombre de un modelo |
| 11 | **Tabla de sensibilidad** | J21:K26 contiene 4 pruebas registradas |
| 12 | **Gráfico comparativo** | 4 series visibles con título y leyenda |
| 13 | **Hoja de recomendación** | Todos los campos completados |
| 14 | **Archivo guardado** | Nombre: `Practica14_PMP_TuNombre.xlsx` |

### Prueba de coherencia de fórmulas

Selecciona manualmente una fila intermedia (por ejemplo, fila 15, correspondiente a S14) y verifica:

```
PMP-A en E15 = 0.2 × B13 + 0.3 × B14 + 0.5 × B15
PMP-B en G15 = 0.1 × B13 + 0.3 × B14 + 0.6 × B15
Error ABS PMP-A en F15 = |B15 - E15|
Error ABS PMP-B en H15 = |B15 - G15|
```

Si alguno de estos valores no coincide con el cálculo manual, revisa las referencias de celda en las fórmulas correspondientes.

---

## 8. Resolución de Problemas

### Problema 1: La suma de pesos muestra un valor diferente de 1 (por ejemplo, 0.99 o 1.01)

**Síntoma:** La celda K5 o K10 muestra un valor como `0.9999999` o `1.0000001`, y el formato condicional la resalta en rojo, aunque visualmente los pesos parecen sumar 1.

**Causa:** Excel almacena números decimales en formato de punto flotante binario, lo que puede generar pequeños errores de redondeo al sumar números como 0.5 + 0.3 + 0.2. Este es un comportamiento conocido del estándar IEEE 754 y no indica un error en los datos ingresados.

**Solución:**
1. Modifica la fórmula de validación en K5 para que use redondeo:
   ```excel
   =REDONDEAR(K2+K3+K4, 10)
   ```
2. Actualiza la regla de formato condicional para que use una tolerancia:
   - Cambia la fórmula de `=K5<>1` por `=ABS(K5-1)>0.0001`
3. Alternativamente, ingresa los pesos como fracciones enteras en celdas separadas y divídelos: por ejemplo, usa valores 5, 3, 2 en celdas auxiliares y divide entre 10 en las celdas K2:K4.

---

### Problema 2: Los pronósticos PMP-A y PMP-B muestran valores idénticos a PMS-3

**Síntoma:** Al comparar las columnas C, E y G, los valores son exactamente iguales (o casi idénticos) en todas las filas. El MAE de los tres modelos también resulta igual.

**Causa:** Las fórmulas en las columnas E y G están referenciando las celdas de parámetros en el orden incorrecto, o están usando una referencia absoluta que apunta a la misma celda para los tres pesos. Esto puede ocurrir si al arrastrar la fórmula se "congelaron" todas las referencias con `$` o si los tres pesos accidentalmente apuntan al mismo valor.

**Solución:**
1. Haz clic en E4 y examina la fórmula en la barra de fórmulas. Debe verse así:
   ```excel
   =$K$4*B2 + $K$3*B3 + $K$2*B4
   ```
   Si ves algo como `=$K$2*B2 + $K$2*B3 + $K$2*B4`, las referencias están mal.
2. Verifica que K2 = 0.5, K3 = 0.3 y K4 = 0.2 (no todos el mismo valor).
3. Reescribe la fórmula manualmente en E4 asegurándote de que cada peso referencia una celda diferente, y vuelve a arrastrar hacia abajo.
4. Para G4, verifica que las referencias sean K7, K8 y K9 (no K2, K3, K4 nuevamente).

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza los siguientes pasos de cierre:

1. **Restaura los pesos del Esquema A** a sus valores originales si modificaste los parámetros durante el ejercicio de sensibilidad y no los restauraste al final del Paso 5:
   - K2: `0.5` | K3: `0.3` | K4: `0.2`
   - Verifica que K5 = 1.

2. **Guarda el archivo final** con `Ctrl + S`. Confirma que el nombre es `Practica14_PMP_TuNombre.xlsx`.

3. **Verifica la estructura del archivo:** el libro debe contener al menos 2 hojas:
   - `Hoja1` (o el nombre original de la Práctica 13): con todas las columnas de datos y la zona de parámetros.
   - `Recomendacion_Operativa`: con la tabla de conclusiones completada.
   - Opcionalmente: `Gráfico_Comparativo` si moviste el gráfico a una hoja separada.

4. **Entrega o comparte el archivo** según las instrucciones del instructor. Este archivo será el insumo para la Práctica 15, donde se aplicará Suavizamiento Exponencial Simple al mismo dataset.

5. **No elimines ninguna columna** del archivo. Las columnas de errores absolutos (D, F, H) y los MAE de la tabla resumen serán referenciados en prácticas posteriores para la comparación final de modelos.

---

## 10. Resumen y Recursos Adicionales

### Lo que aprendiste en esta práctica

En esta práctica implementaste el **Promedio Móvil Ponderado (PMP)**, una evolución directa del PMS que supera una de sus principales limitaciones: el trato igualitario de todos los períodos históricos. Los puntos clave que trabajaste son:

- **La lógica del PMP:** asignar mayor peso al período más reciente permite que el modelo reaccione más rápidamente a cambios en la demanda, lo cual es valioso en entornos logísticos donde las condiciones del mercado cambian con frecuencia.

- **La restricción de suma de pesos = 1:** esta condición no es arbitraria; garantiza que el pronóstico se mantenga en la misma escala que la demanda real. Si los pesos sumaran más de 1, el pronóstico inflaría sistemáticamente las estimaciones.

- **`SUMAPRODUCTO` como herramienta de ponderación:** esta función es extremadamente versátil en Excel para cálculos ponderados, y su dominio será útil en múltiples análisis logísticos más allá del forecasting.

- **El ejercicio de sensibilidad:** modificar los pesos y observar el cambio en el MAE es una versión manual y simplificada de lo que hacen los algoritmos de optimización de parámetros. Desarrollar esta intuición es fundamental para interpretar y validar modelos de pronóstico en la práctica.

- **La conexión operativa:** el pronóstico no es un fin en sí mismo. Su valor está en la decisión que habilita: cuánto pedir, cuándo pedir, cuántos recursos de picking asignar.

### Conexión con el ciclo metodológico del curso

Esta práctica forma parte del ciclo de evaluación de modelos de forecasting logístico:

```
Datos históricos → Depuración (P13) → PMS-3 → PMP (P14) → Suavizamiento Exponencial (P15)
                                                    ↓
                                          Comparación por MAE/MAPE
                                                    ↓
                                        Selección del mejor modelo
                                                    ↓
                                     Decisión operativa de inventario
```

### Recursos adicionales

| Recurso | Descripción | Acceso |
|---|---|---|
| Hyndman & Athanasopoulos — *Forecasting: Principles and Practice* | Capítulo 3: Métodos de suavizamiento. Referencia académica gratuita en línea. | [otexts.com/fpp3](https://otexts.com/fpp3/) |
| Microsoft Support — SUMAPRODUCTO | Documentación oficial de la función con ejemplos | [support.microsoft.com](https://support.microsoft.com/es-es/office/funci%C3%B3n-sumaproducto-16753e75-9f68-4874-94ac-4d2145a2fd2e) |
| Chopra & Meindl — *Supply Chain Management* | Capítulo 7: Demand Forecasting in a Supply Chain | Pearson (biblioteca institucional) |
| Makridakis, Wheelwright & Hyndman — *Forecasting: Methods and Applications* | Referencia clásica sobre métodos de promedio móvil y ponderado | Wiley (biblioteca institucional) |

### Vista previa de la Práctica 15

En la **Práctica 15** aplicarás el método de **Suavizamiento Exponencial Simple (SES)**, que lleva la lógica del PMP a su expresión más elegante: en lugar de asignar pesos fijos a un número finito de períodos, el SES asigna pesos que decaen exponencialmente hacia el pasado, usando todos los datos históricos disponibles con un único parámetro α. Compararás el SES contra los modelos de esta práctica usando el mismo framework de MAE, completando el ciclo de evaluación de los tres métodos fundamentales de suavizamiento.

---

> **Recordatorio del instructor:** Verificar que todos los participantes tengan K5 = 1 y K10 = 1 antes de avanzar al Paso 4. Una suma de pesos incorrecta invalidará todos los cálculos de MAE posteriores. Distribuir el archivo `Checkpoint_P14_Pasos1-3.xlsx` a quienes no hayan completado los primeros tres pasos antes de iniciar el Paso 4.

---
LAB_END---

---

---LAB_START---
LAB_ID: 03-00-03
---MARKDOWN---
# Implementación de Suavizamiento Exponencial Simple

## 1. Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 3 — Métodos Cuantitativos de Pronóstico |
| **Práctica número** | 15 |
| **Archivo de entrada** | `dataset_24semanas_SKUs.xlsx` (generado o recibido en Prácticas 13–14) |
| **Archivo de salida** | `practica15_SES.xlsx` |

---

## 2. Descripción General

En las prácticas anteriores (13 y 14) construiste pronósticos basados en promedio móvil simple y ponderado, donde todos los períodos dentro de la ventana reciben igual peso o pesos fijos. El **Suavizamiento Exponencial Simple (SES)** da un paso más: asigna pesos que decaen exponencialmente hacia el pasado, de modo que los datos más recientes tienen siempre mayor influencia que los más antiguos, sin necesidad de definir una ventana rígida. En esta práctica implementarás el modelo SES desde cero en Excel, compararás tres valores del parámetro **α** y utilizarás **Solver** para encontrar el valor óptimo que minimiza el error de pronóstico, conectando directamente el análisis estadístico con decisiones operativas de reabastecimiento e inventario.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Implementar la fórmula recursiva del SES `F(t+1) = α × D(t) + (1-α) × F(t)` en Excel con referencias mixtas y una celda de parámetro centralizada.
- [ ] Comparar el comportamiento de tres valores de α (0.1, 0.3 y 0.5) sobre el mismo dataset de 24 semanas, interpretando el efecto de cada uno sobre la velocidad de respuesta del modelo.
- [ ] Calcular el MAE para cada valor de α y seleccionar el más adecuado con justificación operativa logística.
- [ ] Usar el complemento **Solver** de Excel para optimizar α y minimizar el MAE de forma automática.
- [ ] Visualizar los tres modelos SES junto a la demanda real en un gráfico de líneas e interpretar las diferencias en términos de reactividad y estabilidad del pronóstico.

---

## 4. Prerrequisitos

### Conocimiento previo requerido

| Requisito | Descripción |
|---|---|
| Teoría SES (Módulo 3.3) | Comprensión de la fórmula recursiva y el rol de α como factor de suavizamiento |
| Prácticas 13 y 14 completadas | Familiaridad con el dataset de 24 semanas y cálculo de MAE en Excel |
| Referencias absolutas en Excel | Uso de `$A$1` para anclar celdas de parámetro en fórmulas |
| Función `ABS()` en Excel | Cálculo de error absoluto período a período |

### Acceso y archivos necesarios

| Recurso | Estado requerido |
|---|---|
| Microsoft Excel 2016 o superior (Microsoft 365 recomendado) | ✅ Instalado y activo |
| Complemento **Solver** habilitado en Excel | ✅ Verificar antes de iniciar (ver Sección 5) |
| Archivo `dataset_24semanas_SKUs.xlsx` | ✅ Disponible en carpeta de trabajo |
| Archivo de checkpoint Práctica 14 (opcional) | ✅ Disponible en caso de no haber completado P14 |

---

## 5. Entorno del Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Core i7 / Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366×768 | 1920×1080 |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado |
| Complemento Solver | Incluido en Excel | Debe habilitarse manualmente si no está activo |
| Navegador web | Chrome / Edge actualizados | Solo para referencia; no se usa activamente en esta práctica |

### Configuración previa: Activar el complemento Solver

Antes de iniciar los pasos del laboratorio, verifica que Solver esté disponible en tu instalación de Excel:

1. Abre Excel y ve a **Archivo → Opciones**.
2. En el panel izquierdo, selecciona **Complementos**.
3. En la parte inferior, asegúrate de que el menú desplegable diga **Complementos de Excel** y haz clic en **Ir…**
4. Marca la casilla **Solver** y haz clic en **Aceptar**.
5. Verifica que aparezca el botón **Solver** en la pestaña **Datos** (grupo **Análisis**).

> ⚠️ **Importante:** Si el complemento Solver no aparece en la lista, es posible que tu instalación de Office sea corporativa con restricciones. Consulta al instructor para obtener una alternativa (tabla de búsqueda manual de α).

---

## 6. Pasos del Laboratorio

---

### Paso 1: Preparar la hoja de trabajo SES

**Objetivo:** Crear una hoja limpia y estructurada con el dataset de 24 semanas, lista para construir los tres modelos SES y la sección de parámetro dinámico.

#### Instrucciones

1. Abre el archivo `dataset_24semanas_SKUs.xlsx` utilizado en las Prácticas 13 y 14.
2. Haz clic derecho sobre la pestaña de la hoja activa y selecciona **Mover o copiar**. Crea una copia de la hoja y renómbrala como **`P15_SES`**.
3. En la hoja `P15_SES`, elimina todas las columnas de cálculo previo (promedios móviles, errores anteriores), dejando únicamente las dos columnas base:
   - **Columna A:** `Período` (valores del 1 al 24)
   - **Columna B:** `Demanda_Real` (las 24 semanas de demanda del SKU seleccionado — usa el mismo SKU trabajado en P13 y P14)
4. En la **fila 1** escribe los siguientes encabezados en las columnas indicadas:

```
A1: Período
B1: Demanda_Real
C1: SES_α0.1
D1: SES_α0.3
E1: SES_α0.5
F1: Error_α0.1
G1: Error_α0.3
H1: Error_α0.5
```

5. Aplica formato de encabezado (negrita, fondo azul oscuro, texto blanco) a la fila 1 para distinguirla visualmente.
6. Reserva el rango **J1:K5** para la sección de parámetro dinámico (la construirás en el Paso 4). Por ahora, escribe en `J1` la etiqueta: `PARÁMETRO DINÁMICO` con fondo amarillo.

#### Salida esperada

Una hoja con 24 filas de datos (filas 2 a 25), encabezados en fila 1, columnas A y B con datos limpios, y columnas C a H vacías listas para recibir fórmulas.

#### Verificación

- Confirma que la columna B contiene exactamente 24 valores numéricos sin celdas vacías ni texto.
- Confirma que las columnas C a H están completamente vacías.
- El rango `J1` debe mostrar la etiqueta `PARÁMETRO DINÁMICO`.

---

### Paso 2: Calcular el valor inicial F₁ (promedio de los primeros 3 períodos)

**Objetivo:** Establecer el valor de arranque del modelo SES de forma consistente para los tres valores de α, usando el promedio de los primeros 3 períodos como inicialización.

#### Instrucciones

1. En la celda **C2** (primera fila de datos, columna SES_α0.1), escribe la fórmula de inicialización:

```excel
=PROMEDIO($B$2:$B$4)
```

2. Copia exactamente la misma fórmula en **D2** y **E2**:

```excel
=PROMEDIO($B$2:$B$4)
```

> 💡 **¿Por qué el mismo valor inicial para los tres modelos?** Para que la comparación entre α sea justa, los tres modelos deben partir del mismo punto. Al usar el promedio de los primeros 3 períodos como F₁, garantizamos que las diferencias observadas en los pronósticos se deban exclusivamente al valor de α y no a distintos puntos de partida.

3. Aplica un formato de fondo **gris claro** a las celdas C2, D2 y E2 para indicar visualmente que son valores de inicialización, no pronósticos calculados con la fórmula recursiva.

4. Anota en una celda de comentario (o en la celda `J3`) el valor resultante del promedio inicial. Debería ser algo similar a:

```
F1 inicial = PROMEDIO(D1:D3) = [valor numérico]
```

#### Salida esperada

Las tres celdas C2, D2 y E2 muestran el mismo valor numérico (el promedio de los primeros 3 períodos de demanda real), con fondo gris claro.

#### Verificación

- Verifica que C2 = D2 = E2.
- El valor debe estar en el rango lógico de la demanda del SKU (no debe ser 0 ni un valor extremo).
- Haz clic en C2 y confirma en la barra de fórmulas que dice `=PROMEDIO($B$2:$B$4)`.

---

### Paso 3: Implementar la fórmula recursiva SES para los tres valores de α

**Objetivo:** Construir los pronósticos período a período usando la fórmula `F(t+1) = α × D(t) + (1-α) × F(t)` con α fijo en cada columna.

#### Instrucciones

**Para la columna SES_α0.1 (columna C):**

1. En la celda **C3**, escribe la fórmula recursiva con α = 0.1:

```excel
=0.1*B2+(1-0.1)*C2
```

Esta fórmula dice: "El pronóstico para el período 3 es el 10% de la demanda real del período 2, más el 90% del pronóstico anterior".

2. Selecciona la celda C3 y arrastra la fórmula hacia abajo hasta la celda **C25** (período 24). Esto generará los pronósticos F₃ hasta F₂₅ (el pronóstico F₂₅ es la proyección para la semana 25, es decir, el pronóstico futuro).

**Para la columna SES_α0.3 (columna D):**

3. En la celda **D3**, escribe la fórmula con α = 0.3:

```excel
=0.3*B2+(1-0.3)*D2
```

4. Arrastra la fórmula desde D3 hasta **D25**.

**Para la columna SES_α0.5 (columna E):**

5. En la celda **E3**, escribe la fórmula con α = 0.5:

```excel
=0.5*B2+(1-0.5)*E2
```

6. Arrastra la fórmula desde E3 hasta **E25**.

> ⚠️ **Atención con las referencias:** En la fórmula `=0.1*B2+(1-0.1)*C2`, la referencia a `B2` debe ser relativa (no anclada) para que al arrastrar la fórmula hacia abajo, cada fila tome la demanda real del período correspondiente. La referencia a `C2` también debe ser relativa para tomar el pronóstico del período anterior. **No uses `$` en estas referencias.**

7. Una vez completadas las tres columnas, revisa visualmente que los valores cambien período a período y que no haya valores en blanco ni errores `#¡REF!` o `#¡VALOR!`.

#### Salida esperada

Las columnas C, D y E contienen 24 valores de pronóstico cada una (filas 2 a 25). Los valores de la fila 2 son idénticos en las tres columnas (valor inicial). A partir de la fila 3, los valores divergen según el α de cada columna. La columna E (α=0.5) debe mostrar mayor variabilidad y seguir más de cerca a la demanda real; la columna C (α=0.1) debe mostrar una curva más suave y con mayor rezago.

#### Verificación

Comprueba manualmente el cálculo del período 3 en las tres columnas:

```
C3 = 0.1 × B2 + 0.9 × C2   → debe coincidir con el valor en C3
D3 = 0.3 × B2 + 0.7 × D2   → debe coincidir con el valor en D3
E3 = 0.5 × B2 + 0.5 × E2   → debe coincidir con el valor en E3
```

Realiza este cálculo en papel o en una celda auxiliar y compara con el resultado en Excel. Si hay discrepancia, revisa las referencias de la fórmula.

---

### Paso 4: Crear la sección de parámetro dinámico (celda de α referenciada)

**Objetivo:** Implementar una versión del modelo SES donde α se define en una sola celda y toda la columna la referencia, permitiendo cambiar el valor en un solo lugar y observar el impacto en tiempo real.

#### Instrucciones

1. En el rango reservado **J1:K5**, construye la siguiente tabla de parámetro:

```
J1: === PARÁMETRO DINÁMICO ===     (fondo amarillo, negrita)
J2: Valor de α:
K2: 0.3                            (este es el valor editable — fondo amarillo intenso)
J3: F1 inicial:
K3: =PROMEDIO($B$2:$B$4)          (mismo valor inicial que los modelos anteriores)
J4: MAE modelo dinámico:
K4: (se completará en el Paso 5)
```

2. Asigna el nombre `alfa_dinamico` a la celda **K2** para facilitar su uso en fórmulas:
   - Selecciona K2 → Ve al **Cuadro de nombres** (a la izquierda de la barra de fórmulas) → Escribe `alfa_dinamico` → Presiona Enter.

3. En la columna **I**, agrega el encabezado `SES_Dinámico` en **I1**.

4. En la celda **I2**, escribe el valor inicial (igual que los modelos fijos):

```excel
=$K$3
```

5. En la celda **I3**, escribe la fórmula recursiva usando la referencia al parámetro dinámico:

```excel
=$K$2*B2+(1-$K$2)*I2
```

O usando el nombre definido:

```excel
=alfa_dinamico*B2+(1-alfa_dinamico)*I2
```

6. Arrastra la fórmula desde **I3** hasta **I25**.

7. **Prueba el comportamiento dinámico:** Cambia el valor en K2 de 0.3 a 0.7 y observa cómo todos los valores de la columna I se recalculan instantáneamente. Luego, vuelve a colocar 0.3 para continuar con el siguiente paso.

> 💡 **Relevancia logística:** Esta celda de parámetro simula lo que haría un analista de demanda al ajustar la sensibilidad del modelo según el comportamiento del producto. Un α alto conviene para productos con demanda muy volátil (ej. productos de moda o temporada); un α bajo es mejor para productos de demanda estable (ej. commodities de consumo masivo).

#### Salida esperada

La columna I contiene 24 valores de pronóstico que se recalculan automáticamente al modificar el valor en K2. Al cambiar K2, los valores de I3:I25 y el MAE en K4 (una vez calculado) deben actualizarse sin necesidad de modificar ninguna otra celda.

#### Verificación

- Cambia K2 a 0.1 → los valores de I deben coincidir aproximadamente con los de la columna C.
- Cambia K2 a 0.5 → los valores de I deben coincidir aproximadamente con los de la columna E.
- Restaura K2 = 0.3 antes de continuar.

---

### Paso 5: Calcular errores absolutos y MAE para cada modelo

**Objetivo:** Medir la precisión de cada modelo SES calculando el error absoluto período a período y el MAE total, para comparar objetivamente los tres valores de α.

#### Instrucciones

**Cálculo de errores absolutos (columnas F, G, H):**

> 📌 **Nota sobre el período de evaluación:** El error se calcula comparando el pronóstico del período `t` con la demanda real del período `t`. Dado que el pronóstico `F(t)` se genera en el período `t-1`, el primer error calculable es en el período 2 (comparando F₂ con D₂). Sin embargo, como F₂ es el valor inicial (promedio de los primeros 3 períodos), muchos analistas omiten los primeros períodos del cálculo del MAE. En esta práctica, calcularemos el error desde el **período 4** (fila 5) hasta el **período 24** (fila 25), excluyendo los 3 períodos de inicialización.

1. En la celda **F5**, escribe la fórmula de error absoluto para α = 0.1:

```excel
=ABS(B5-C4)
```

Esta fórmula compara la demanda real del período 4 (`B5`) con el pronóstico que el modelo α=0.1 generó para ese período (`C4`).

2. Arrastra la fórmula desde F5 hasta **F25**.

3. En la celda **G5**, escribe para α = 0.3:

```excel
=ABS(B5-D4)
```

4. Arrastra desde G5 hasta **G25**.

5. En la celda **H5**, escribe para α = 0.5:

```excel
=ABS(B5-E4)
```

6. Arrastra desde H5 hasta **H25**.

**Cálculo del MAE total:**

7. En una zona de resumen (sugerida: rango **J7:K10**), construye la siguiente tabla:

```
J7:  === RESUMEN MAE ===           (negrita, fondo azul claro)
J8:  MAE  α = 0.1
K8:  =PROMEDIO(F5:F25)
J9:  MAE  α = 0.3
K9:  =PROMEDIO(G5:G25)
J10: MAE  α = 0.5
K10: =PROMEDIO(H5:H25)
```

8. Aplica formato de número con 2 decimales a las celdas K8, K9 y K10.

9. Agrega formato condicional al rango K8:K10 para resaltar en **verde** el valor mínimo (el mejor modelo):
   - Selecciona K8:K10 → **Inicio → Formato condicional → Reglas superiores/inferiores → Valores mínimos** → Selecciona formato verde.

10. Actualiza la celda **K4** (MAE modelo dinámico) con la fórmula:

```excel
=PROMEDIO(ABS(B5-I4),ABS(B6-I5),ABS(B7-I6),ABS(B8-I7))
```

> ⚠️ Esta fórmula es solo una aproximación para K4. Opcionalmente, puedes crear una columna de error para el modelo dinámico similar a las columnas F-H. Para simplicidad, usa:

```excel
=PROMEDIO(MAPA(B5:B25,I4:I24,LAMBDA(d,f,ABS(d-f))))
```

O la alternativa más compatible con Excel 2016:

```excel
=SUMPRODUCT(ABS(B5:B25-I4:I24))/21
```

#### Salida esperada

Las columnas F, G y H contienen 21 valores de error absoluto cada una (períodos 4 al 24). La tabla de resumen en J7:K10 muestra los tres valores de MAE. El valor mínimo aparece resaltado en verde, indicando el modelo con mejor desempeño predictivo para este dataset.

#### Verificación

- Los errores en las columnas F, G y H deben ser todos valores positivos (la función `ABS` garantiza esto).
- Verifica manualmente el error del período 5:
  ```
  F5 = |B5 - C4|   → comprueba con calculadora
  ```
- El MAE de α=0.3 debe ser menor que el de α=0.1 y α=0.5 en la mayoría de datasets con demanda moderadamente estable (aunque el resultado depende de los datos específicos).

---

### Paso 6: Crear el gráfico de líneas comparativo

**Objetivo:** Visualizar simultáneamente la demanda real y los tres modelos SES para interpretar el comportamiento de cada valor de α de forma gráfica.

#### Instrucciones

1. Selecciona el rango **A1:E25** (columnas Período, Demanda_Real, SES_α0.1, SES_α0.3, SES_α0.5).

2. Ve a **Insertar → Gráficos → Gráfico de líneas → Líneas con marcadores**.

3. Excel generará un gráfico con 4 series. Ajusta el formato de cada serie para facilitar la comparación:

| Serie | Color sugerido | Grosor de línea | Marcadores |
|---|---|---|---|
| Demanda_Real | Negro | 2.5 pt | Círculo sólido |
| SES_α0.1 | Azul claro | 1.5 pt | Sin marcadores |
| SES_α0.3 | Naranja | 1.5 pt | Sin marcadores |
| SES_α0.5 | Rojo | 1.5 pt | Sin marcadores |

4. Agrega los siguientes elementos al gráfico:
   - **Título:** `Comparación de Modelos SES — α = 0.1, 0.3 y 0.5`
   - **Título del eje X:** `Período (Semanas)`
   - **Título del eje Y:** `Unidades`
   - **Leyenda:** Posición inferior

5. Mueve el gráfico a una posición que no tape los datos (debajo de la fila 27 o en una hoja separada llamada `Gráfico_SES`).

6. Observa el gráfico y responde mentalmente (o anota en la celda `J13`) las siguientes preguntas de interpretación:
   - ¿Qué modelo sigue más de cerca a la demanda real?
   - ¿Qué modelo produce la curva más suave?
   - ¿En qué períodos el modelo α=0.1 muestra mayor rezago respecto a la demanda real?

> 💡 **Conexión con el contenido teórico:** Este gráfico ilustra visualmente el mismo principio discutido en la Lección 3.1 sobre el efecto del tamaño de ventana N en el promedio móvil: un α pequeño en SES es análogo a un N grande en promedio móvil (más suavizado, más rezago); un α grande es análogo a un N pequeño (más reactivo, más ruido).

#### Salida esperada

Un gráfico de líneas con 4 series claramente diferenciadas por color. La línea de demanda real (negra) muestra la variabilidad original; las líneas SES muestran distintos grados de suavizamiento, siendo la de α=0.5 la más cercana a la demanda real y la de α=0.1 la más suave.

#### Verificación

- Confirma que el gráfico tiene exactamente 4 series (no más, no menos).
- Verifica que el eje X muestre los 24 períodos correctamente etiquetados.
- La línea de α=0.5 debe cruzarse más frecuentemente con la línea de demanda real que las otras dos.

---

### Paso 7: Optimizar α con Solver

**Objetivo:** Usar el complemento Solver de Excel para encontrar automáticamente el valor de α que minimiza el MAE del modelo dinámico, completando el ciclo de evaluación y selección de modelos.

#### Instrucciones

**Preparación previa:**

1. Antes de ejecutar Solver, asegúrate de que:
   - La celda **K2** contiene el valor inicial `0.3` (punto de partida para la optimización).
   - La celda **K4** contiene la fórmula del MAE del modelo dinámico (calculada en el Paso 5).
   - La columna I contiene los pronósticos del modelo dinámico referenciados a K2.

**Configuración de Solver:**

2. Ve a la pestaña **Datos** → haz clic en **Solver** (grupo Análisis).

3. En el cuadro de diálogo de Solver, configura los siguientes parámetros:

| Campo de Solver | Valor a ingresar |
|---|---|
| **Establecer objetivo** | `$K$4` (la celda del MAE dinámico) |
| **Para** | Selecciona **Mínimo** |
| **Cambiando las celdas variables** | `$K$2` (la celda del parámetro α) |

4. Agrega las siguientes restricciones haciendo clic en **Agregar**:

   - **Restricción 1:** `$K$2 >= 0.01` (α debe ser mayor que 0 para que el modelo tenga sentido)
   - **Restricción 2:** `$K$2 <= 0.99` (α debe ser menor que 1 para mantener memoria del pasado)

5. En el campo **Método de resolución**, selecciona **GRG Nonlinear** (apropiado para funciones no lineales continuas como la minimización del MAE en SES).

6. Haz clic en **Resolver**.

7. Cuando Solver encuentre la solución, aparecerá un cuadro de diálogo indicando "Solver encontró una solución". Selecciona **Mantener solución de Solver** y haz clic en **Aceptar**.

**Documentación del resultado:**

8. En el rango **J15:K19**, documenta los resultados de la optimización:

```
J15: === RESULTADO SOLVER ===      (negrita, fondo verde claro)
J16: α óptimo encontrado:
K16: [valor que Solver colocó en K2]
J17: MAE con α óptimo:
K17: =K4
J18: MAE con α = 0.3:
K18: =K9
J19: Mejora relativa (%):
K19: =(K18-K17)/K18*100
```

9. Aplica formato de porcentaje con 1 decimal a K19.

10. Anota una conclusión operativa en la celda **J21** (texto libre, máximo 2 líneas). Ejemplo:

```
"El α óptimo de [valor] indica que para este SKU conviene dar un peso 
[alto/moderado/bajo] a la demanda reciente. Esto implica [mayor/menor] 
sensibilidad ante picos de demanda, lo cual [favorece/dificulta] la 
planificación de reabastecimiento en un entorno de demanda [estable/volátil]."
```

#### Salida esperada

La celda K2 muestra el valor de α óptimo encontrado por Solver (típicamente entre 0.1 y 0.7 dependiendo del dataset). La celda K4 muestra el MAE mínimo alcanzable. La tabla J15:K19 documenta la mejora obtenida respecto al α = 0.3 utilizado como punto de partida.

#### Verificación

- El valor en K2 debe estar entre 0.01 y 0.99 (dentro de las restricciones definidas).
- El MAE en K17 debe ser **menor o igual** al MAE de cualquiera de los tres modelos fijos (K8, K9, K10). Si es mayor, hay un error en la configuración de Solver — revisa que la celda objetivo sea K4 y que el modelo dinámico esté correctamente referenciado a K2.
- La mejora relativa en K19 debe ser un valor positivo (aunque puede ser pequeña si α=0.3 ya era cercano al óptimo).

---

## 7. Validación y Pruebas Finales

Una vez completados todos los pasos, realiza las siguientes verificaciones de integridad del modelo:

### Lista de verificación de integridad

| Verificación | Cómo comprobarlo | ✅ |
|---|---|---|
| Los tres modelos SES tienen el mismo F₁ inicial | C2 = D2 = E2 | ☐ |
| Las fórmulas recursivas de C3:E25 no contienen referencias absolutas innecesarias | Revisar barra de fórmulas en C3, D3, E3 | ☐ |
| Los errores en columnas F, G, H son todos ≥ 0 | `=MIN(F5:H25)` debe ser ≥ 0 | ☐ |
| El MAE del modelo dinámico (K4) coincide con el MAE del α activo en K2 | Pon K2 = 0.1 → K4 debe ≈ K8; pon K2 = 0.3 → K4 debe ≈ K9 | ☐ |
| El gráfico muestra exactamente 4 series con leyenda correcta | Verificar visualmente | ☐ |
| El α óptimo de Solver produce el MAE más bajo entre todos los modelos | K17 ≤ MIN(K8, K9, K10) | ☐ |
| La conclusión operativa en J21 está redactada en términos logísticos | Revisar contenido de J21 | ☐ |

### Prueba de consistencia cruzada

Ejecuta la siguiente verificación para confirmar que el modelo dinámico replica correctamente los modelos fijos:

```excel
' En una celda auxiliar, verifica que el modelo dinámico con α=0.1 
' produce resultados iguales al modelo fijo de la columna C:
' Paso 1: Coloca K2 = 0.1
' Paso 2: En una celda auxiliar escribe: =SUMA(ABS(C3:C25-I3:I25))
' Resultado esperado: 0 (o muy cercano a 0, diferencias por redondeo)
' Confirma y luego restaura K2 al valor óptimo de Solver
```

---

## 8. Solución de Problemas

### Problema 1: Las fórmulas recursivas muestran el mismo valor en todas las filas (no hay variación entre períodos)

**Síntoma:** Al arrastrar la fórmula de C3 hacia abajo, todos los valores de C3:C25 son idénticos o extremadamente similares, sin reflejar los cambios en la demanda real.

**Causa probable:** Las referencias a `B2` y `C2` en la fórmula recursiva fueron accidentalmente ancladas con `$`, impidiendo que se desplacen al arrastrar. Por ejemplo, la fórmula dice `=0.1*$B$2+(1-0.1)*$C$2` en lugar de `=0.1*B2+(1-0.1)*C2`.

**Solución:**
1. Haz clic en la celda C3 y revisa la barra de fórmulas.
2. Si ves signos `$` en las referencias de `B` o `C`, elimínalos manualmente.
3. La fórmula correcta en C3 es: `=0.1*B2+(1-0.1)*C2` (sin ningún `$`).
4. Elimina el contenido de C4:C25 y vuelve a arrastrar la fórmula corregida desde C3.
5. Repite el proceso para las columnas D y E si el problema también aparece ahí.

> ⚠️ **Excepción:** La celda C2 (valor inicial) sí debe tener referencias absolutas: `=PROMEDIO($B$2:$B$4)`. Solo las fórmulas recursivas de C3 en adelante deben tener referencias relativas.

---

### Problema 2: Solver no mejora el MAE o devuelve un α fuera del rango esperado (ej. 0.99 o 0.01)

**Síntoma:** Después de ejecutar Solver, el valor en K2 es exactamente 0.99 o 0.01 (los límites de las restricciones), y el MAE en K4 no mejora respecto a los modelos fijos, o incluso empeora.

**Causa probable A — Celda objetivo incorrecta:** Solver está optimizando una celda que no está vinculada al modelo dinámico de la columna I. Esto ocurre si K4 contiene una fórmula estática (un número escrito directamente) o si hace referencia a la columna C, D o E en lugar de la columna I.

**Causa probable B — Modelo dinámico desconectado:** Las fórmulas de la columna I no referencian a K2 correctamente. Si las fórmulas usan `0.3` como número fijo en lugar de `$K$2`, cambiar K2 no afecta la columna I ni el MAE.

**Solución:**
1. Haz clic en I3 y verifica en la barra de fórmulas que la fórmula sea `=$K$2*B2+(1-$K$2)*I2` (con `$K$2` anclado).
2. Cambia manualmente K2 a 0.9 y observa si los valores de la columna I cambian. Si no cambian, las fórmulas de la columna I no están referenciando K2.
3. Verifica que K4 contenga una fórmula que calcule el MAE basándose en los valores actuales de la columna I (no una fórmula estática ni referenciada a otra columna).
4. Corrige las fórmulas según corresponda, restaura K2 = 0.3 y vuelve a ejecutar Solver.
5. Si el problema persiste, prueba cambiando el método de resolución en Solver de **GRG Nonlinear** a **Evolutionary** y vuelve a ejecutar.

---

## 9. Limpieza y Cierre

### Guardar el archivo de trabajo

1. Guarda el archivo con el nombre **`practica15_SES_[TuNombre].xlsx`** en tu carpeta de trabajo del curso.
2. Asegúrate de que la hoja `P15_SES` sea la hoja activa al guardar, para que el instructor pueda revisar el trabajo directamente al abrir el archivo.
3. Si el archivo tiene otras hojas de prácticas anteriores, **no las elimines** — son insumos para prácticas futuras.

### Verificación de archivos generados

| Archivo | Contenido | ¿Guardado? |
|---|---|---|
| `practica15_SES_[TuNombre].xlsx` | Hoja `P15_SES` con modelos SES, errores, MAE, gráfico y resultado Solver | ☐ |

### Restaurar el parámetro dinámico

Antes de cerrar, asegúrate de que la celda **K2** muestra el **valor óptimo encontrado por Solver** (no 0.3 ni ningún otro valor de prueba). Esto garantiza que el archivo quede en su estado final óptimo para revisión.

### Nota para la siguiente práctica

El archivo `practica15_SES_[TuNombre].xlsx` será referenciado en la **Práctica 16**, donde compararás los resultados del SES con los modelos de promedio móvil de las Prácticas 13 y 14, consolidando la selección del modelo óptimo para cada SKU del portafolio.

---

## 10. Resumen y Recursos Adicionales

### Lo que aprendiste en esta práctica

En esta práctica implementaste el **Suavizamiento Exponencial Simple (SES)** de principio a fin en Excel, abarcando todo el ciclo de un analista de pronóstico:

1. **Inicialización:** Estableciste un punto de partida consistente para los tres modelos usando el promedio de los primeros 3 períodos.
2. **Construcción recursiva:** Aplicaste la fórmula `F(t+1) = α × D(t) + (1-α) × F(t)` con tres valores de α distintos, comprendiendo que α controla el balance entre reactividad y estabilidad del pronóstico.
3. **Parametrización dinámica:** Centralizaste el parámetro α en una sola celda, una práctica de buena ingeniería de modelos que facilita el mantenimiento y la exploración de escenarios.
4. **Evaluación cuantitativa:** Calculaste el MAE para cada modelo y usaste formato condicional para identificar visualmente el mejor desempeño.
5. **Optimización automática:** Empleaste el complemento Solver para encontrar el α óptimo matemáticamente, cerrando el ciclo con una decisión basada en datos.
6. **Interpretación logística:** Conectaste el valor de α con decisiones operativas reales: un α alto implica mayor reactividad ante cambios de demanda (útil en productos volátiles), mientras que un α bajo produce pronósticos más estables (adecuado para productos de demanda predecible).

### Comparación conceptual: SES vs. Promedio Móvil

| Dimensión | Promedio Móvil (N fijo) | SES (α variable) |
|---|---|---|
| Pesos asignados | Iguales para todos los N períodos | Decrecientes exponencialmente hacia el pasado |
| Parámetro clave | Tamaño de ventana N | Factor de suavizamiento α |
| Reactividad | Controlada por N (menor N = más reactivo) | Controlada por α (mayor α = más reactivo) |
| Datos requeridos | Solo los últimos N períodos | Todo el historial (implícitamente en F(t)) |
| Implementación en Excel | `PROMEDIO()` sobre rango deslizante | Fórmula recursiva con referencia a celda anterior |
| Optimización del parámetro | Prueba y error con distintos N | Solver puede encontrar α óptimo automáticamente |

### Recursos adicionales

- **Hyndman, R.J. & Athanasopoulos, G.** — *Forecasting: Principles and Practice*, Capítulo 8: Exponential Smoothing. Disponible gratuitamente en: [https://otexts.com/fpp3/ses.html](https://otexts.com/fpp3/ses.html)
- **Microsoft Support** — Uso del complemento Solver en Excel: [https://support.microsoft.com/es-es/office/definir-y-resolver-un-problema-con-solver-5d1a388f-079d-43ac-a7eb-f63e45925040](https://support.microsoft.com/es-es/office/definir-y-resolver-un-problema-con-solver-5d1a388f-079d-43ac-a7eb-f63e45925040)
- **Chopra, S. & Meindl, P.** — *Supply Chain Management*, Capítulo 7: Demand Forecasting in a Supply Chain (Pearson).
- **ASCM/APICS** — Glosario de términos de gestión de la cadena de suministro: [https://www.ascm.org](https://www.ascm.org)

---

> 🔗 **Conexión con la siguiente práctica:** En la **Práctica 16** consolidarás los resultados de los tres métodos trabajados (Promedio Móvil, Promedio Móvil Ponderado y SES) en una tabla comparativa de MAE por SKU, seleccionando el modelo recomendado para cada producto y justificando la decisión en términos operativos de inventario y nivel de servicio.

---
LAB_END---

---

# Práctica 16 — Creación de gráficos comparativos entre demanda real y pronóstico generado

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Crear (*Create*) |
| **Módulo** | 3 — Métodos Cuantitativos de Pronóstico |
| **Práctica número** | 16 (de 16 en la secuencia) |
| **Archivos requeridos** | Libro Excel con resultados de Prácticas 13, 14 y 15 (hojas: `PMS`, `PMP`, `SES`) |

---

## 2. Descripción General

En esta práctica el estudiante consolidará en un único panel visual los resultados de los tres métodos de pronóstico trabajados en las Prácticas 13, 14 y 15: Promedio Móvil Simple (PMS), Promedio Móvil Ponderado (PMP) y Suavizamiento Exponencial Simple (SES). Se construirá una hoja llamada `Dashboard` que integre un gráfico de líneas comparativo con cinco series, un gráfico de barras de error absoluto por período y una tabla de métricas MAE vinculada dinámicamente. El objetivo final es que el estudiante pueda leer los gráficos y emitir una recomendación operativa justificada sobre cuál método de pronóstico conviene utilizar para el producto analizado, conectando el análisis visual directamente con decisiones de inventario y reabastecimiento.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica serás capaz de:

- [ ] **Diseñar** una hoja `Dashboard` en Excel que integre en un solo panel los gráficos comparativos de demanda real versus pronósticos PMS, PMP y SES, siguiendo buenas prácticas de visualización de datos.
- [ ] **Construir** un gráfico de líneas con al menos cinco series (demanda real, PMS-3, PMP esquema óptimo, SES α óptimo y SES α=0.1) con formato profesional: título, leyenda, etiquetas de ejes y colores diferenciados.
- [ ] **Crear** un gráfico de barras de error absoluto por período para cada método, identificando visualmente los intervalos de mayor desviación.
- [ ] **Vincular dinámicamente** una tabla de métricas MAE a las celdas de cálculo de las hojas de origen, de manera que los valores se actualicen automáticamente.
- [ ] **Interpretar** los resultados visuales y redactar una recomendación operativa justificada de 5–7 líneas sobre el método de pronóstico más adecuado para el producto.

---

## 4. Prerrequisitos

### Conocimientos previos

| Área | Nivel requerido |
|---|---|
| Promedio Móvil Simple (PMS) — Lección 3.1 | Completado |
| Promedio Móvil Ponderado (PMP) — Lección 3.2 | Completado |
| Suavizamiento Exponencial Simple (SES) — Lección 3.3 | Completado |
| Prácticas 13, 14 y 15 con modelos de pronóstico construidos | Completado |
| Inserción y formato básico de gráficos en Excel (series, ejes, leyendas) | Básico–Intermedio |
| Métricas de error MAE y MAPE calculadas en prácticas anteriores | Comprensión conceptual |

### Acceso y archivos

- Libro Excel de trabajo con las hojas `PMS`, `PMP` y `SES` completas y guardadas.
- Acceso a Microsoft Excel 2016 o superior (Microsoft 365 recomendado).
- Si el estudiante no completó alguna práctica anterior, el instructor proporcionará el archivo *checkpoint* correspondiente antes de iniciar.

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Espacio en disco libre | 500 MB | 2 GB |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior | Microsoft 365 recomendado para funciones dinámicas |
| Microsoft Edge / Google Chrome | Versión 2024 actualizada | Solo si se accede a Copilot como apoyo |

### Preparación del entorno

Antes de iniciar, verifica lo siguiente:

1. Abre el libro Excel de trabajo (archivo generado en Prácticas 13–15).
2. Confirma que existen las hojas: `PMS`, `PMP` y `SES`.
3. Verifica que cada hoja contiene las columnas de: período, demanda real, valores de pronóstico y error absoluto calculados.
4. Guarda una copia de seguridad del archivo con el nombre `Practica16_backup.xlsx` antes de comenzar.

```
Nombre sugerido del archivo de trabajo: Pronosticos_SKU_Practicas13a16.xlsx
```

---

## 6. Desarrollo Paso a Paso

---

### Paso 1 — Crear la hoja Dashboard y estructurar la tabla de consolidación

**Objetivo:** Centralizar en una nueva hoja todos los datos necesarios para construir los gráficos comparativos, vinculando dinámicamente las celdas de origen.

#### Instrucciones

1. En el libro Excel de trabajo, haz clic derecho sobre cualquier pestaña de hoja y selecciona **Insertar → Hoja de cálculo**. Renombra la nueva hoja como `Dashboard`.

2. En la hoja `Dashboard`, a partir de la celda **A1**, construye la siguiente tabla de consolidación de datos. Las celdas de las columnas B a F deben ser **referencias dinámicas** a las hojas de origen, no valores copiados:

   | Celda | Encabezado |
   |---|---|
   | A1 | `Período` |
   | B1 | `Demanda Real` |
   | C1 | `PMS-3` |
   | D1 | `PMP (Óptimo)` |
   | E1 | `SES (α óptimo)` |
   | F1 | `SES (α=0.1)` |

3. En la columna **A** (filas 2 en adelante), ingresa los identificadores de período tal como aparecen en tus hojas de origen (ejemplo: `Mes 1`, `Mes 2`, … `Mes 12`, o los nombres de mes reales). Deben coincidir exactamente con los períodos para los cuales existen pronósticos calculados.

   > **Nota:** Los primeros N−1 períodos del PMS no tienen pronóstico (son celdas vacías en la hoja `PMS`). Incluye todos los períodos para los que al menos un método tenga valor calculado.

4. En la columna **B** (Demanda Real), ingresa referencias a la columna de demanda real de la hoja `PMS` (o la hoja que contenga el dato original). Ejemplo para la fila 2:

   ```excel
   =PMS!B2
   ```

   Arrastra esta fórmula hacia abajo para todos los períodos.

5. En la columna **C** (PMS-3), referencia la columna de pronóstico PMS de la hoja correspondiente:

   ```excel
   =PMS!C2
   ```

   Arrastra hacia abajo. Las celdas sin pronóstico quedarán en blanco o mostrarán `#N/A` si usaste `N/A()` en la práctica anterior; ambas situaciones son aceptables para la visualización.

6. En la columna **D** (PMP Óptimo), referencia la columna del mejor esquema de ponderación calculado en la Práctica 14:

   ```excel
   =PMP!D2
   ```

   Ajusta la letra de columna según tu estructura real en la hoja `PMP`.

7. En la columna **E** (SES α óptimo), referencia la columna de pronóstico SES con el alfa que produjo menor MAE en la Práctica 15:

   ```excel
   =SES!E2
   ```

8. En la columna **F** (SES α=0.1), referencia la columna de pronóstico SES con α=0.1:

   ```excel
   =SES!C2
   ```

   Ajusta según tu estructura en la hoja `SES`.

9. Guarda el archivo (**Ctrl + S**).

#### Resultado esperado

La hoja `Dashboard` debe mostrar una tabla limpia con 5 columnas de datos numéricos (más la columna de período), donde todas las celdas de las columnas B a F son fórmulas de referencia que muestran los mismos valores que las hojas de origen. Si modificas un valor en `PMS`, el cambio debe reflejarse automáticamente en `Dashboard`.

#### Verificación

- Haz clic en cualquier celda de la columna B (por ejemplo, `B3`) y confirma en la barra de fórmulas que el contenido es una referencia del tipo `=PMS!B3`, no un valor estático.
- Verifica que el número de filas con datos en `Dashboard` coincide exactamente con el número de períodos con pronóstico en las hojas de origen.

---

### Paso 2 — Construir el gráfico de líneas comparativo (5 series)

**Objetivo:** Crear un gráfico de líneas profesional que muestre simultáneamente la demanda real y los cuatro pronósticos, permitiendo comparación visual inmediata entre métodos.

#### Instrucciones

1. En la hoja `Dashboard`, selecciona el rango completo de la tabla de consolidación creada en el Paso 1. Por ejemplo, si tus datos van de la fila 1 a la fila 13 (12 meses), selecciona **A1:F13**.

2. Ve a la pestaña **Insertar** → grupo **Gráficos** → selecciona **Gráfico de líneas** → elige **Línea con marcadores** (*Line with Markers*). Excel insertará el gráfico con las 5 series automáticamente.

3. **Reubica el gráfico:** Haz clic en el borde del gráfico y arrástralo para que quede posicionado aproximadamente en el rango **H1:R20** de la hoja `Dashboard`, dejando espacio a la izquierda para la tabla de datos y debajo para el segundo gráfico.

4. **Asigna un título descriptivo:** Haz doble clic sobre el título del gráfico y escribe:

   ```
   Comparativo de Demanda Real vs. Métodos de Pronóstico (PMS, PMP, SES)
   ```

5. **Configura el eje horizontal (X):** Haz clic derecho sobre el gráfico → **Seleccionar datos** (*Select Data*). En la sección **Etiquetas del eje horizontal**, haz clic en **Editar** y selecciona el rango de la columna A (períodos) de tu tabla. Esto garantiza que el eje muestre los nombres de período correctos.

6. **Aplica colores diferenciados por serie.** Haz clic derecho sobre cada línea del gráfico → **Dar formato a serie de datos** y asigna los siguientes colores recomendados (puedes ajustar según preferencia, pero deben ser claramente distintos):

   | Serie | Color recomendado | Grosor de línea |
   |---|---|---|
   | Demanda Real | Negro (`#000000`) | 2.5 pt |
   | PMS-3 | Azul (`#1F77B4`) | 1.5 pt |
   | PMP (Óptimo) | Naranja (`#FF7F0E`) | 1.5 pt |
   | SES (α óptimo) | Verde (`#2CA02C`) | 1.5 pt |
   | SES (α=0.1) | Rojo punteado (`#D62728`) | 1.5 pt, estilo guión |

   > **Consejo:** Para cambiar el estilo de línea a guión, en el panel de formato busca **Tipo de guión** y selecciona la opción de línea discontinua.

7. **Etiqueta los ejes:**
   - Haz clic en el gráfico → botón **+** (Elementos del gráfico) → activa **Títulos de ejes**.
   - Eje vertical (Y): escribe `Unidades (cajas)`
   - Eje horizontal (X): escribe `Período`

8. **Configura la leyenda:** Asegúrate de que la leyenda esté visible. Haz clic en el botón **+** → **Leyenda** → selecciona posición **Abajo** para no interferir con las líneas del gráfico.

9. **Ajusta el rango del eje Y:** Haz clic derecho sobre el eje Y → **Dar formato al eje**. Establece un mínimo manual que sea aproximadamente el 80% del valor mínimo de tu demanda real (esto evita que las diferencias entre métodos se vean exageradas o aplastadas). Por ejemplo, si tu demanda mínima es 400, establece el mínimo del eje en 300.

10. **Aplica un estilo de gráfico limpio:** En la pestaña **Diseño de gráfico** (que aparece al seleccionar el gráfico), selecciona el **Estilo 3** o equivalente con fondo blanco y líneas de cuadrícula tenues. Elimina el fondo gris si lo hubiera.

11. Guarda el archivo (**Ctrl + S**).

#### Resultado esperado

El gráfico de líneas debe mostrar claramente las 5 series con colores distintos, título descriptivo, leyenda identificada y ejes etiquetados con unidades. Las líneas de pronóstico deben comenzar en el período donde cada método genera su primer valor (el PMS comenzará más tarde que el SES, lo cual es visualmente correcto y esperado).

#### Verificación

- Confirma que el gráfico muestra exactamente 5 series en la leyenda.
- Verifica que el eje X muestra los identificadores de período (no números del 1 al N).
- Comprueba que la serie "Demanda Real" es visualmente distinguible del resto (línea más gruesa o color negro).

---

### Paso 3 — Crear el gráfico de barras de error absoluto por período

**Objetivo:** Construir un segundo gráfico que muestre el error absoluto de cada método en cada período, facilitando la identificación visual de los intervalos donde cada modelo falla más.

#### Instrucciones

1. En la hoja `Dashboard`, debajo de la tabla de consolidación (por ejemplo, a partir de la fila 16 si tus datos terminan en la fila 13), crea una segunda tabla con los errores absolutos. Usa los siguientes encabezados a partir de la celda **A16**:

   | Celda | Encabezado |
   |---|---|
   | A16 | `Período` |
   | B16 | `Error Abs. PMS-3` |
   | C16 | `Error Abs. PMP` |
   | D16 | `Error Abs. SES (α óptimo)` |
   | E16 | `Error Abs. SES (α=0.1)` |

2. En la columna **A** (fila 17 en adelante), copia o referencia los mismos identificadores de período de la tabla principal.

3. En la columna **B** (Error Abs. PMS-3), referencia la columna de error absoluto calculada en la hoja `PMS`. Si en esa hoja el error absoluto está en la columna D, la fórmula sería:

   ```excel
   =PMS!D2
   ```

   Arrastra hacia abajo para todos los períodos.

4. Repite el mismo procedimiento para las columnas C, D y E, referenciando las columnas de error absoluto de las hojas `PMP` y `SES` respectivamente.

   > **Importante:** Los períodos sin pronóstico (primeros N−1 del PMS) tendrán error absoluto de 0 o celda vacía. Para que el gráfico sea más limpio, usa la siguiente fórmula condicional que deja la celda en blanco si no hay pronóstico:
   >
   > ```excel
   > =SI(PMS!C2="","",ABS(PMS!B2-PMS!C2))
   > ```

5. Selecciona el rango completo de la segunda tabla (por ejemplo, **A16:E28** si tienes 12 períodos).

6. Ve a **Insertar** → **Gráficos** → **Gráfico de barras** → selecciona **Barras agrupadas** (*Clustered Bar*) o, preferiblemente, **Columnas agrupadas** (*Clustered Column*) para que el eje del tiempo quede horizontal.

7. Reubica este segundo gráfico debajo del primero, aproximadamente en el rango **H22:R38** de la hoja `Dashboard`.

8. **Asigna título:**

   ```
   Error Absoluto por Período — Comparativo de Métodos de Pronóstico
   ```

9. **Etiqueta los ejes:**
   - Eje Y: `Error Absoluto (unidades)`
   - Eje X: `Período`

10. **Aplica los mismos colores** usados en el gráfico de líneas para cada método, de manera que el estudiante pueda relacionar visualmente ambos gráficos sin necesidad de leer la leyenda.

11. Activa la leyenda en posición **Abajo**.

12. Guarda el archivo (**Ctrl + S**).

#### Resultado esperado

El gráfico de columnas agrupadas debe mostrar para cada período cuatro barras de colores distintos (una por método), permitiendo identificar de un vistazo en qué períodos el error fue mayor y qué método tuvo el peor desempeño en esos momentos. Los períodos sin pronóstico del PMS aparecerán sin barra o con barra en cero.

#### Verificación

- Confirma que el gráfico muestra 4 series (una por método, sin incluir "Demanda Real" que no tiene error absoluto propio).
- Verifica que los colores de las barras coinciden con los colores de las líneas del gráfico anterior para el mismo método.
- Comprueba que el eje X muestra los identificadores de período correctos.

---

### Paso 4 — Crear la tabla de métricas MAE vinculada dinámicamente

**Objetivo:** Construir una tabla resumen de métricas de error (MAE) que se actualice automáticamente al modificar los datos de origen, integrando el análisis cuantitativo con la visualización.

#### Instrucciones

1. En la hoja `Dashboard`, selecciona la celda **A30** (o una celda libre debajo de las tablas anteriores). Construye la siguiente tabla de métricas:

   ```
   A30: TABLA DE MÉTRICAS — RESUMEN DE ERROR
   ```

   Aplica **negrita** y color de relleno azul oscuro con texto blanco para esta celda de título. Combina las celdas A30:C30 para el título.

2. En las filas siguientes, construye la tabla de métricas:

   | Celda | Contenido |
   |---|---|
   | A31 | `Método` |
   | B31 | `MAE` |
   | C31 | `Interpretación` |
   | A32 | `PMS-3` |
   | A33 | `PMP (Óptimo)` |
   | A34 | `SES (α óptimo)` |
   | A35 | `SES (α=0.1)` |

3. En la celda **B32**, ingresa una referencia dinámica a la celda donde calculaste el MAE del PMS en la hoja `PMS`. Por ejemplo:

   ```excel
   =PMS!F15
   ```

   Ajusta la referencia de celda según donde esté el MAE calculado en tu hoja `PMS`. **No copies el valor manualmente; debe ser una referencia de fórmula.**

4. Repite para las celdas B33, B34 y B35, referenciando los MAE de `PMP` y `SES` respectivamente:

   ```excel
   B33: =PMP!F15
   B34: =SES!G15
   B35: =SES!C15
   ```

   Ajusta según la ubicación real de cada MAE en las hojas correspondientes.

5. En la columna **C** (Interpretación), ingresa fórmulas condicionales que clasifiquen automáticamente cada método:

   ```excel
   C32: =SI(B32=MIN(B32:B35),"✔ Menor error","")
   C33: =SI(B33=MIN(B32:B35),"✔ Menor error","")
   C34: =SI(B34=MIN(B32:B35),"✔ Menor error","")
   C35: =SI(B35=MIN(B32:B35),"✔ Menor error","")
   ```

   Esto marcará automáticamente el método con menor MAE.

6. **Aplica formato condicional** a las celdas B32:B35:
   - Selecciona B32:B35 → **Inicio** → **Formato condicional** → **Escalas de color**.
   - Elige una escala de **verde a rojo** (verde = menor MAE, rojo = mayor MAE). Esto permite identificar visualmente el mejor método.

7. **Aplica bordes** a toda la tabla (A31:C35): selecciona el rango → **Inicio** → **Bordes** → **Todos los bordes**.

8. Aplica **negrita** a la fila de encabezados (A31:C31) y color de fondo gris claro.

9. Guarda el archivo (**Ctrl + S**).

#### Resultado esperado

La tabla de métricas debe mostrar los cuatro valores MAE con formato numérico de 2 decimales, con formato condicional de escala de color que resalta el mejor método en verde, y la columna de interpretación marcando automáticamente el método con menor error con el símbolo ✔.

#### Verificación

- Modifica temporalmente un valor de pronóstico en la hoja `PMS` (por ejemplo, cambia un valor por 999) y verifica que el MAE en la celda B32 del `Dashboard` se actualiza automáticamente. Luego deshaz el cambio (**Ctrl + Z**).
- Confirma que la fórmula `MIN` en la columna C identifica correctamente el método con menor MAE.

---

### Paso 5 — Añadir cuadro de texto con interpretación y recomendación operativa

**Objetivo:** Redactar e insertar una interpretación analítica de 5–7 líneas que conecte los hallazgos visuales con una recomendación operativa concreta para decisiones de inventario y reabastecimiento.

#### Instrucciones

1. En la hoja `Dashboard`, ve a la pestaña **Insertar** → grupo **Texto** → selecciona **Cuadro de texto** (*Text Box*).

2. Dibuja el cuadro de texto debajo de la tabla de métricas, aproximadamente en el rango **A37:G50**, con un tamaño suficiente para contener el texto sin que quede apretado.

3. Dentro del cuadro de texto, redacta tu interpretación siguiendo esta estructura de 5–7 líneas. El texto debe basarse en los valores reales de MAE que obtuviste en tus prácticas. A continuación se muestra una **plantilla de referencia** que debes personalizar con tus datos reales:

   ```
   INTERPRETACIÓN Y RECOMENDACIÓN OPERATIVA

   Análisis comparativo de métodos de pronóstico — [Nombre del SKU]
   Período analizado: [Mes inicio] a [Mes fin] | Total: [N] períodos

   Con base en los gráficos comparativos y las métricas de error calculadas, el método
   [NOMBRE DEL MÉTODO] presenta el menor MAE ([valor]), lo que indica que sus pronósticos
   se desvían en promedio [valor] unidades respecto a la demanda real. Visualmente, este
   método sigue de forma más cercana la línea de demanda real, especialmente en los períodos
   [identificar períodos de buen desempeño].

   El método PMS-3 muestra un rezago visible en los períodos de cambio de tendencia (períodos
   [X] a [Y]), lo cual es consistente con su limitación estructural: al promediar los últimos
   3 períodos con igual peso, reacciona tarde ante variaciones en la demanda. Esto implicaría
   un riesgo de subabastecimiento si la demanda aumenta de forma sostenida.

   RECOMENDACIÓN: Para este producto, se recomienda utilizar [MÉTODO] como modelo base de
   pronóstico para el ciclo de reabastecimiento. Este método permitirá estimar con mayor
   precisión el volumen de órdenes de compra y el stock de seguridad necesario. Se sugiere
   revisar el modelo cada [período] para ajustar parámetros si el comportamiento de la
   demanda cambia significativamente.
   ```

4. **Personaliza el texto** con los valores reales de tu análisis: nombre del SKU, períodos analizados, valores de MAE, nombre del método ganador y observaciones específicas sobre los gráficos.

5. **Aplica formato al cuadro de texto:**
   - Borde: línea sólida azul oscuro, grosor 1.5 pt.
   - Relleno: color de fondo azul muy claro (`#EBF5FB`) o gris claro.
   - Fuente del título interno: **negrita**, tamaño 11.
   - Fuente del cuerpo: tamaño 10, color negro.

6. **Añade una etiqueta de título** sobre el cuadro de texto. Puedes hacerlo insertando un segundo cuadro de texto pequeño o usando la primera línea del cuadro principal con formato de título:

   ```
   📋 INTERPRETACIÓN Y RECOMENDACIÓN OPERATIVA
   ```

7. Guarda el archivo (**Ctrl + S**).

#### Resultado esperado

El cuadro de texto debe mostrar una interpretación coherente y personalizada de 5–7 líneas que: (a) identifica el método con menor MAE, (b) señala al menos un período de mayor error visible en los gráficos, (c) conecta la limitación del PMS con el concepto de rezago aprendido en la Lección 3.1, y (d) emite una recomendación operativa concreta vinculada a decisiones de inventario o reabastecimiento.

#### Verificación

- Lee el texto en voz alta y verifica que menciona valores numéricos reales (no genéricos).
- Confirma que la recomendación hace referencia explícita al impacto en una decisión logística (orden de compra, stock de seguridad, nivel de servicio o planificación de picking).

---

### Paso 6 — Aplicar formato final y organizar el Dashboard

**Objetivo:** Dar coherencia visual al panel completo, asegurando que todos los elementos estén alineados, etiquetados y con un aspecto profesional listo para presentación.

#### Instrucciones

1. **Añade un encabezado de Dashboard** en la fila 1 de la hoja `Dashboard` (antes de la tabla de datos). Inserta una fila vacía al inicio (**clic derecho sobre fila 1 → Insertar**) y en la celda **A1** escribe:

   ```
   DASHBOARD DE ANÁLISIS DE PRONÓSTICO — [NOMBRE DEL SKU] — [FECHA]
   ```

   Combina las celdas A1:R1 (**Inicio → Combinar y centrar**), aplica fondo azul marino, texto blanco, negrita, tamaño 14.

2. **Verifica la organización espacial** del Dashboard. El layout recomendado es:

   ```
   Fila 1:        [ENCABEZADO PRINCIPAL — combinado A1:R1]
   Filas 2-14:    [Tabla de consolidación A:F] | [Gráfico de líneas H:R]
   Filas 15-28:   [Tabla de errores A:E]       | [Gráfico de barras H:R]
   Filas 29-50:   [Tabla MAE A:C]              | [Cuadro de texto A:G]
   ```

3. **Congela la fila de encabezado:** Ve a **Vista** → **Inmovilizar paneles** → **Inmovilizar fila superior**. Esto asegura que el encabezado siempre sea visible al desplazarse.

4. **Oculta las líneas de cuadrícula** de la hoja para dar aspecto más limpio: **Vista** → desactiva la casilla **Líneas de cuadrícula** (*Gridlines*).

5. **Verifica que los dos gráficos estén alineados:** Selecciona ambos gráficos manteniendo **Ctrl** presionado → clic derecho → **Alinear** → **Alinear a la izquierda**. Esto los deja con el mismo margen izquierdo.

6. **Añade un pie de página informativo** (opcional pero recomendado): Ve a **Insertar** → **Encabezado y pie de página** → en el pie de página escribe:

   ```
   Práctica 16 — Análisis Comparativo de Pronósticos | Generado: [fecha]
   ```

7. Realiza una última revisión visual completa del Dashboard. Verifica:
   - Todos los gráficos tienen título.
   - Todos los ejes tienen etiquetas con unidades.
   - La tabla MAE tiene formato condicional activo.
   - El cuadro de interpretación tiene texto personalizado (no la plantilla genérica).

8. Guarda el archivo con un nombre final descriptivo:

   ```
   Practica16_Dashboard_Pronosticos_[TusIniciales].xlsx
   ```

#### Resultado esperado

La hoja `Dashboard` debe presentar un panel visual completo, profesional y autocontenido que un gerente de logística pueda interpretar sin necesidad de acceder a las hojas de cálculo intermedias. Todos los elementos deben estar alineados, con colores consistentes entre gráficos y tabla de métricas.

#### Verificación

- Cierra el archivo y vuelve a abrirlo. Verifica que todos los gráficos muestran los datos correctamente (a veces los vínculos se rompen si el archivo se mueve de carpeta).
- Imprime o exporta la hoja `Dashboard` como PDF (**Archivo → Exportar → Crear PDF**) y verifica que todos los elementos caben en una o dos páginas con orientación horizontal.

---

## 7. Validación y Pruebas Finales

Una vez completados todos los pasos, realiza las siguientes verificaciones de integridad del Dashboard:

### Lista de verificación de entregable

| # | Elemento a verificar | ✅ / ❌ |
|---|---|---|
| 1 | La hoja `Dashboard` existe y contiene todos los elementos | |
| 2 | La tabla de consolidación tiene 5 columnas de datos vinculadas dinámicamente (no valores estáticos) | |
| 3 | El gráfico de líneas muestra exactamente 5 series con colores diferenciados | |
| 4 | El gráfico de líneas tiene título descriptivo, leyenda y etiquetas de ejes con unidades | |
| 5 | El gráfico de barras muestra 4 series de error absoluto (una por método) | |
| 6 | El gráfico de barras tiene título, leyenda y etiquetas de ejes | |
| 7 | Los colores de las barras coinciden con los colores de las líneas para el mismo método | |
| 8 | La tabla MAE tiene 4 filas (una por método) con referencias dinámicas a hojas de origen | |
| 9 | La tabla MAE tiene formato condicional de escala de color activo | |
| 10 | La columna de interpretación marca automáticamente el método con menor MAE | |
| 11 | El cuadro de texto tiene interpretación personalizada (no la plantilla genérica) | |
| 12 | La interpretación menciona al menos un valor numérico real de MAE | |
| 13 | La recomendación hace referencia a una decisión logística concreta | |
| 14 | El Dashboard tiene encabezado principal con nombre del SKU y fecha | |
| 15 | El archivo se guarda con el nombre correcto incluyendo las iniciales del estudiante | |

### Prueba de integridad dinámica

Para confirmar que los vínculos dinámicos funcionan correctamente:

1. Ve a la hoja `PMS` y modifica temporalmente el valor de demanda real del período 5 (cambia el valor original por `999`).
2. Regresa a la hoja `Dashboard` y verifica que:
   - La celda correspondiente en la columna "Demanda Real" de la tabla de consolidación muestra `999`.
   - El gráfico de líneas refleja el cambio visual (la línea negra sube abruptamente en el período 5).
   - El MAE del PMS en la tabla de métricas ha cambiado.
3. Deshaz el cambio con **Ctrl + Z** dos veces y confirma que todos los valores regresan a su estado original.

---

## 8. Resolución de Problemas

### Problema 1 — El gráfico de líneas no muestra el eje X con los nombres de período correctos (muestra 1, 2, 3... en lugar de "Mes 1", "Mes 2"...)

**Síntoma:** El eje horizontal del gráfico de líneas muestra números secuenciales (1, 2, 3…) en lugar de los identificadores de período definidos en la columna A de la tabla de consolidación.

**Causa:** Al insertar el gráfico, Excel interpreta la columna A como una serie de datos numérica adicional en lugar de como etiquetas del eje horizontal. Esto ocurre cuando la columna A contiene valores como "1", "2", "3" (numéricos) o cuando el rango de selección no incluye correctamente la columna A como categoría.

**Solución:**
1. Haz clic derecho sobre el gráfico → **Seleccionar datos** (*Select Data*).
2. En el panel izquierdo, verifica si aparece "Período" como una serie de datos. Si es así, selecciónala y haz clic en **Quitar** (*Remove*).
3. En el panel derecho (**Etiquetas del eje horizontal**), haz clic en **Editar** (*Edit*).
4. En el campo "Rango de etiquetas del eje", selecciona manualmente el rango de la columna A que contiene los identificadores de período (por ejemplo, `Dashboard!$A$2:$A$13`).
5. Haz clic en **Aceptar** en ambos paneles.
6. El eje X ahora debe mostrar los nombres de período correctos.

---

### Problema 2 — Las referencias dinámicas de la tabla MAE muestran `#¡REF!` o un valor incorrecto

**Síntoma:** Las celdas B32:B35 de la tabla MAE muestran el error `#¡REF!` o un valor que no coincide con el MAE calculado en las hojas de origen (`PMS`, `PMP`, `SES`).

**Causa:** La referencia de celda en la fórmula apunta a una celda incorrecta en la hoja de origen. Esto ocurre frecuentemente cuando: (a) el estudiante calculó el MAE en una celda diferente a la indicada en el enunciado, (b) se insertaron o eliminaron filas en las hojas de origen después de crear la referencia, o (c) el nombre de la hoja de origen fue modificado (por ejemplo, de `PMS` a `Practica13`).

**Solución:**
1. Ve a la hoja de origen correspondiente (por ejemplo, `PMS`) y localiza visualmente la celda que contiene el valor MAE calculado. Anota su dirección exacta (por ejemplo, `F15`).
2. Regresa a la hoja `Dashboard` y haz clic en la celda con el error (por ejemplo, `B32`).
3. Edita la fórmula directamente en la barra de fórmulas. Reemplaza la referencia incorrecta por la dirección correcta:
   ```excel
   =PMS!F15
   ```
   Si el nombre de la hoja fue cambiado, actualiza también el nombre:
   ```excel
   =Practica13!F15
   ```
4. Si las filas fueron insertadas o eliminadas en la hoja de origen y ahora el MAE está en una fila diferente, actualiza el número de fila en la referencia.
5. Presiona **Enter** y verifica que el valor mostrado coincide con el MAE visible en la hoja de origen.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones de cierre:

1. **Guarda el archivo final** con el nombre definitivo:
   ```
   Practica16_Dashboard_Pronosticos_[TusIniciales].xlsx
   ```

2. **Exporta el Dashboard como PDF** para conservar una copia de presentación independiente del archivo Excel:
   - Ve a **Archivo** → **Exportar** → **Crear documento PDF/XPS**.
   - En opciones de publicación, selecciona **Solo hoja activa** y orientación **Horizontal**.
   - Guarda el PDF con el nombre: `Practica16_Dashboard_[TusIniciales].pdf`

3. **Verifica que las hojas de origen (`PMS`, `PMP`, `SES`) no fueron modificadas** durante la práctica. Si algún valor fue alterado durante las pruebas de verificación, asegúrate de haberlo revertido con **Ctrl + Z**.

4. **Cierra Excel** correctamente guardando todos los cambios.

5. Si el instructor solicita entrega digital, sube ambos archivos (`.xlsx` y `.pdf`) a la plataforma indicada antes de finalizar la sesión.

---

## 10. Resumen y Recursos Adicionales

### Resumen de la Práctica

En esta práctica construiste un **Dashboard de análisis de pronóstico** completamente funcional en Excel que integra los resultados de tres métodos cuantitativos (PMS, PMP y SES) en un único panel visual. Los elementos clave que desarrollaste fueron:

| Elemento | Propósito logístico |
|---|---|
| **Gráfico de líneas (5 series)** | Comparar visualmente el seguimiento de cada método respecto a la demanda real a lo largo del tiempo |
| **Gráfico de barras de error absoluto** | Identificar los períodos donde cada método falla más, informando decisiones de ajuste del modelo |
| **Tabla MAE con vínculos dinámicos** | Cuantificar y comparar el desempeño de cada método con actualización automática |
| **Cuadro de interpretación** | Traducir el análisis estadístico en una recomendación operativa concreta para el equipo de logística |

La habilidad central desarrollada en esta práctica —construir visualizaciones que conectan el análisis cuantitativo con decisiones operativas— es directamente aplicable en roles de planificación de demanda, gestión de inventarios y coordinación de operaciones de picking. Recuerda que **un pronóstico no tiene valor si no puede ser comunicado e interpretado por quienes toman decisiones**.

### Conexión con la Lección 3.1 (Promedio Móvil Simple)

El gráfico de líneas construido en esta práctica hace evidente uno de los conceptos fundamentales de la Lección 3.1: el **rezago estructural del promedio móvil**. Al visualizar la serie PMS-3 junto a la demanda real, puedes observar que la línea de pronóstico siempre "llega tarde" a los cambios de nivel, especialmente en períodos de crecimiento o caída sostenida. Este comportamiento —que en la lección se describió como una limitación del método— ahora es cuantificable mediante el MAE y visible en el gráfico comparativo, completando el ciclo de aprendizaje: concepto → cálculo → visualización → decisión.

### Recursos Adicionales

| Recurso | Descripción | Acceso |
|---|---|---|
| **Microsoft Support — Crear un gráfico de principio a fin** | Guía oficial para insertar y formatear gráficos en Excel | [support.microsoft.com](https://support.microsoft.com/es-es/office/crear-un-gr%C3%A1fico-de-principio-a-fin-0baf399e-dd61-4e18-8a73-b3fd5d5680c2) |
| **Hyndman & Athanasopoulos — Forecasting: Principles and Practice** | Capítulo 3: Time series decomposition y métodos de suavizamiento | [otexts.com/fpp3](https://otexts.com/fpp3/) |
| **Makridakis et al. — Forecasting: Methods and Applications** | Referencia clásica sobre métricas de error y comparación de modelos | Wiley, 3.ª edición |
| **Chopra & Meindl — Supply Chain Management** | Capítulo sobre gestión de la demanda y pronóstico en cadena de suministro | Pearson, edición vigente |

---

> **Nota para el instructor:** Esta práctica cierra el ciclo de las Prácticas 13–16 y constituye el entregable integrador del módulo de métodos cuantitativos. Se recomienda dedicar los últimos 3–4 minutos de la sesión a una revisión rápida en pantalla compartida donde 2–3 estudiantes muestren su Dashboard y lean su recomendación operativa en voz alta, promoviendo la discusión sobre por qué diferentes estudiantes pueden llegar a recomendaciones distintas según los datos de su SKU asignado.

---

---LAB_START---
LAB_ID: 03-00-05
---MARKDOWN---
# Práctica 17 — Generación de Pronóstico Automático con Herramienta Integrada de Excel

## 1. Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 3 — Métodos de Pronóstico Cuantitativo |
| **Práctica número** | 17 |
| **Archivo de entrada** | `Lab_17_Demanda_NuevoProducto.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Lab_17_Pronostico_Completo.xlsx` |

---

## 2. Descripción General

En esta práctica utilizarás la herramienta **Hoja de Pronóstico** integrada en Excel para generar automáticamente un pronóstico de demanda con intervalos de confianza al 95%, a partir de un historial de 18 meses de un nuevo producto. Complementarás este análisis aplicando manualmente el método de **Suavizamiento Exponencial Simple (SES)** sobre el mismo dataset y comparando ambos resultados. Finalmente, consultarás una herramienta de **IA Generativa** para investigar el método estadístico subyacente (ETS — *Error, Trend, Seasonality*) que Excel utiliza internamente, enriqueciendo tu comprensión conceptual del pronóstico automático. Esta práctica conecta directamente con los métodos manuales de las Prácticas 13–15 y extiende la lógica del promedio móvil aprendida en la Lección 3.1 hacia herramientas más sofisticadas.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] Generar un pronóstico automático de 6 meses con intervalos de confianza al 95% utilizando la herramienta **Hoja de Pronóstico** de Excel (pestaña Datos).
- [ ] Interpretar las columnas de pronóstico, límite inferior y límite superior generadas automáticamente, y construir un gráfico con área sombreada para el intervalo de confianza.
- [ ] Aplicar SES manualmente sobre el mismo dataset y comparar cuantitativamente sus proyecciones contra el pronóstico automático de Excel.
- [ ] Documentar los parámetros que Excel seleccionó automáticamente (método ETS, valor de alfa, detección de estacionalidad) y explicar su significado operativo.
- [ ] Consultar a una herramienta de IA Generativa sobre el método ETS y documentar la respuesta para vincularla con decisiones de inventario y nivel de servicio.

---

## 4. Prerrequisitos

### Conocimiento previo requerido

| Tema | Dónde se cubrió |
|---|---|
| Cálculo de Promedio Móvil Simple (PMS) | Lección 3.1 / Práctica 13 |
| Promedio Móvil Ponderado (PMP) | Lección 3.2 / Práctica 14 |
| Suavizamiento Exponencial Simple (SES) | Lección 3.3 / Práctica 15 |
| Métricas de error MAE y MAPE | Lección 3.4 / Práctica 16 |
| Uso básico de gráficos en Excel | Módulo 2 |

### Acceso y archivos necesarios

- [ ] Archivo `Lab_17_Demanda_NuevoProducto.xlsx` descargado y guardado localmente.
- [ ] Excel 2016 o superior instalado y funcional (verificar que la pestaña **Datos** muestre el botón **Hoja de Pronóstico**).
- [ ] Acceso a una herramienta de IA Generativa: Microsoft Copilot (`copilot.microsoft.com`), ChatGPT (`chat.openai.com`) o equivalente.
- [ ] Conexión a internet activa (mínimo 10 Mbps).

---

## 5. Entorno de Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | i7 / Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Conexión a internet | 10 Mbps | 25 Mbps |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | La herramienta **Hoja de Pronóstico** no está disponible en versiones anteriores a 2016 |
| Navegador web | Chrome / Edge (versión 2024) | Para acceso a IA Generativa |
| Herramienta de IA Generativa | Copilot, ChatGPT o Gemini | Acceso gratuito o licencia estándar |

### Verificación de entorno antes de comenzar

Antes de iniciar los pasos del laboratorio, ejecuta estas verificaciones rápidas:

**Verificación 1 — Hoja de Pronóstico disponible en Excel:**
1. Abre Excel y crea un libro en blanco temporal.
2. Ve a la pestaña **Datos**.
3. Confirma que el grupo **Pronóstico** contiene el botón **Hoja de Pronóstico** (ícono de gráfico de líneas con proyección).
4. Si no aparece: actualiza Excel a través de **Archivo → Cuenta → Opciones de actualización → Actualizar ahora**.

**Verificación 2 — Acceso a IA Generativa:**
1. Abre tu navegador y navega a `copilot.microsoft.com` o `chat.openai.com`.
2. Confirma que puedes escribir y recibir respuestas.
3. Si no tienes acceso, solicita al instructor las capturas de pantalla de referencia preparadas para esta práctica.

**Verificación 3 — Archivo de datos:**
1. Abre `Lab_17_Demanda_NuevoProducto.xlsx`.
2. Confirma que contiene dos columnas: **Fecha** (columna A, formato mes/año, 18 filas de datos) y **Demanda_Real** (columna B, valores numéricos enteros).
3. Confirma que no hay celdas vacías en el rango de datos.

---

## 6. Pasos del Laboratorio

> **Nota de tiempo:** Esta práctica tiene 20 minutos. La distribución sugerida es: Pasos 1–3 (8 min) → Pasos 4–5 (6 min) → Paso 6 (4 min) → Paso 7 (2 min). Si te retrasas en el Paso 3, solicita al instructor el archivo de *checkpoint* `Lab_17_CP1_HojaPronostico.xlsx`.

---

### Paso 1 — Exploración del Dataset y Preparación del Archivo

**Objetivo:** Familiarizarte con los datos de demanda del nuevo producto y preparar el archivo para el análisis de pronóstico.

#### Instrucciones

1. Abre el archivo `Lab_17_Demanda_NuevoProducto.xlsx`. Deberías ver la siguiente estructura en la hoja **Datos_Historicos**:

   | Columna A | Columna B |
   |---|---|
   | Fecha | Demanda_Real |
   | ene-23 | (valor numérico) |
   | feb-23 | (valor numérico) |
   | … | … |
   | jun-24 | (valor numérico) |

2. Selecciona el rango **A1:B19** (encabezados + 18 meses de datos) y aplica formato de tabla:
   - Ve a **Inicio → Dar formato como tabla** → elige un estilo con encabezado.
   - Confirma que Excel detecta los encabezados correctamente.

3. En la celda **D1**, escribe el encabezado `Análisis Exploratorio`. Completa las siguientes celdas con fórmulas estadísticas básicas para caracterizar la demanda antes de pronosticar:

   ```excel
   D2: Promedio
   E2: =PROMEDIO(B2:B19)

   D3: Mediana
   E3: =MEDIANA(B2:B19)

   D4: Desv. Estándar
   E4: =DESVEST.M(B2:B19)

   D5: Coef. Variación
   E5: =E4/E2

   D6: Mínimo
   E6: =MIN(B2:B19)

   D7: Máximo
   E7: =MAX(B2:B19)
   ```

4. Formatea la celda **E5** como porcentaje con 1 decimal.

5. Guarda el archivo con el nombre `Lab_17_Pronostico_Completo.xlsx` usando **Archivo → Guardar como**.

#### Resultado esperado

La tabla de análisis exploratorio muestra valores calculados para los 18 meses. El coeficiente de variación (CV) te da una primera señal sobre la estabilidad de la demanda: si CV < 0.20, la demanda es estable; si CV > 0.40, es volátil. Anota mentalmente este valor porque lo usarás para interpretar la calidad del pronóstico en el Paso 5.

#### Verificación

✅ El archivo `Lab_17_Pronostico_Completo.xlsx` está guardado con los datos originales intactos en la hoja **Datos_Historicos** y las estadísticas descriptivas en las celdas D2:E7.

---

### Paso 2 — Generación del Pronóstico Automático con Hoja de Pronóstico

**Objetivo:** Utilizar la herramienta nativa de Excel para generar automáticamente un pronóstico de 6 meses con intervalo de confianza al 95%.

#### Instrucciones

1. Asegúrate de que el cursor esté dentro del rango de datos **A1:B19** (haz clic en cualquier celda dentro del rango de datos).

2. Ve a la pestaña **Datos** → grupo **Pronóstico** → haz clic en **Hoja de Pronóstico**.

   > Si el botón no aparece, confirma que estás en Excel 2016 o superior. En versiones anteriores a 2016 esta herramienta no existe.

3. Se abrirá el cuadro de diálogo **Crear hoja de cálculo de pronóstico**. Configura los siguientes parámetros:

   | Parámetro | Valor a configurar |
   |---|---|
   | **Fin del pronóstico** | `dic-24` (6 meses después del último dato, jun-24) |
   | **Intervalo de confianza** | `95%` (verificar que esté en 95, no en 80 o 90) |
   | **Estacionalidad** | Selecciona **Detectar automáticamente** en primera instancia |
   | **Tipo de gráfico** | Gráfico de líneas (opción por defecto, mantenerla) |

4. Antes de hacer clic en **Crear**, haz clic en el enlace **Opciones** (parte inferior izquierda del cuadro de diálogo) para expandir las opciones avanzadas. Observa y **documenta en papel o en una celda de notas** los siguientes campos:

   - **Inicio del pronóstico:** ¿Qué fecha aparece?
   - **Completar puntos usando:** ¿Qué método aparece seleccionado?
   - **Agregar estadísticas:** Activa esta casilla si está disponible.

5. Haz clic en **Crear**.

6. Excel generará automáticamente una **nueva hoja** llamada `Hoja de pronóstico1` (o similar). Esta hoja contiene:
   - Los datos históricos originales.
   - Las columnas de pronóstico generadas automáticamente.
   - Un gráfico integrado con el área sombreada del intervalo de confianza.

7. Renombra esta hoja como `Pronostico_Excel_Auto` haciendo doble clic en la pestaña.

#### Resultado esperado

La hoja `Pronostico_Excel_Auto` contiene una tabla con columnas similares a las siguientes (los nombres exactos pueden variar según la versión de Excel):

| Fecha | Demanda_Real | Pronóstico | Límite inferior (95%) | Límite superior (95%) |
|---|---|---|---|---|
| ene-23 | 245 | 245 | — | — |
| … | … | … | … | … |
| jun-24 | 312 | 312 | — | — |
| jul-24 | — | ~320 | ~285 | ~355 |
| ago-24 | — | ~318 | ~278 | ~358 |
| … | — | … | … | … |
| dic-24 | — | ~330 | ~270 | ~390 |

El gráfico integrado muestra la línea histórica en azul sólido, la línea de pronóstico en azul discontinuo y el área sombreada representando el intervalo de confianza al 95%.

#### Verificación

✅ La hoja `Pronostico_Excel_Auto` existe y contiene exactamente 6 filas de pronóstico futuro (jul-24 a dic-24) con valores en las tres columnas: Pronóstico, Límite inferior y Límite superior.

✅ El gráfico integrado muestra el área sombreada del intervalo de confianza.

---

### Paso 3 — Documentación de Parámetros Automáticos de Excel

**Objetivo:** Registrar los parámetros que Excel seleccionó automáticamente para comprender qué método está usando internamente.

#### Instrucciones

1. En el libro `Lab_17_Pronostico_Completo.xlsx`, crea una nueva hoja llamada `Parametros_Documentacion`.

2. En esta hoja, construye la siguiente tabla de documentación (escribe los encabezados y completa los valores observados):

   ```
   Celda A1: DOCUMENTACIÓN DE PARÁMETROS — HOJA DE PRONÓSTICO EXCEL
   
   Celda A3: Parámetro
   Celda B3: Valor observado / Seleccionado por Excel
   Celda C3: Interpretación logística
   
   Celda A4:  Método base utilizado
   Celda B4:  [Escribe aquí lo que observaste: ETS, suavizamiento, etc.]
   Celda C4:  [¿Qué significa para el pronóstico de inventario?]
   
   Celda A5:  Detección de estacionalidad
   Celda B5:  [¿Detectó estacionalidad? ¿Cuántos períodos?]
   Celda C5:  [¿Afecta el ciclo de reabastecimiento?]
   
   Celda A6:  Intervalo de confianza configurado
   Celda B6:  95%
   Celda C6:  [¿Qué implica para el stock de seguridad?]
   
   Celda A7:  Amplitud del intervalo en mes 1 (jul-24)
   Celda B7:  [Límite superior - Límite inferior del primer mes]
   Celda C7:  [¿Es aceptable para planificación operativa?]
   
   Celda A8:  Amplitud del intervalo en mes 6 (dic-24)
   Celda B8:  [Límite superior - Límite inferior del sexto mes]
   Celda C8:  [¿Cómo crece la incertidumbre con el horizonte?]
   ```

3. Para calcular la amplitud del intervalo en las celdas B7 y B8, regresa a la hoja `Pronostico_Excel_Auto` y resta manualmente: `Límite superior - Límite inferior` para jul-24 y dic-24. Escribe esos valores en B7 y B8.

4. Observa si la amplitud del intervalo **crece** entre el mes 1 y el mes 6 del pronóstico. Anota tu observación en la celda **A10**:

   ```
   A10: OBSERVACIÓN: La amplitud del intervalo de confianza [crece / se mantiene estable] 
        entre jul-24 y dic-24, lo que indica que la incertidumbre del pronóstico 
        [aumenta / es constante] con el horizonte temporal.
   ```

5. Guarda el archivo.

#### Resultado esperado

La hoja `Parametros_Documentacion` contiene una tabla completa con los parámetros registrados. Es normal que la amplitud del intervalo de confianza **aumente** a medida que el horizonte de pronóstico se extiende, porque la incertidumbre acumulada crece con el tiempo — esto tiene implicaciones directas para el dimensionamiento del stock de seguridad en logística.

#### Verificación

✅ La tabla de parámetros tiene valores registrados en todas las celdas de la columna B (no hay celdas vacías).

✅ La amplitud del intervalo en dic-24 es **mayor** que en jul-24 (si esto no ocurre, revisa que hayas configurado correctamente el intervalo de confianza al 95%).

---

### Paso 4 — Construcción del Gráfico con Área Sombreada para el Intervalo de Confianza

**Objetivo:** Crear un gráfico profesional que visualice el pronóstico y el intervalo de confianza con área sombreada, exportando los datos de la hoja automática a una hoja de trabajo propia.

#### Instrucciones

1. Crea una nueva hoja en el libro y nómbrala `Grafico_Intervalo`.

2. Copia los datos de la hoja `Pronostico_Excel_Auto` hacia esta nueva hoja. Selecciona únicamente las filas correspondientes a los **6 meses futuros** (jul-24 a dic-24) más los **últimos 6 meses históricos** (ene-24 a jun-24) para que el gráfico muestre la transición entre historia y pronóstico. La tabla debe tener esta estructura:

   | A | B | C | D | E |
   |---|---|---|---|---|
   | Fecha | Demanda_Real | Pronóstico | Lim_Inferior | Lim_Superior |
   | ene-24 | (valor) | (valor) | — | — |
   | … | … | … | … | … |
   | jun-24 | (valor) | (valor) | — | — |
   | jul-24 | — | (valor) | (valor) | (valor) |
   | … | — | … | … | … |
   | dic-24 | — | (valor) | (valor) | (valor) |

   > Para las celdas donde no hay valor (por ejemplo, Demanda_Real en meses futuros), deja la celda **vacía** (no escribas cero, porque distorsionaría el gráfico).

3. Selecciona el rango completo de datos de la tabla (A1 hasta E13 aproximadamente).

4. Inserta un **Gráfico de Líneas** básico: **Insertar → Gráficos → Líneas → Líneas con marcadores**.

5. Ahora convierte el área del intervalo de confianza en un área sombreada usando el siguiente procedimiento:

   a. Haz clic derecho sobre la serie **Lim_Inferior** en el gráfico → **Cambiar tipo de gráfico de la serie**.
   b. Cambia **Lim_Inferior** a tipo **Área**.
   c. Cambia **Lim_Superior** a tipo **Área apilada** (esto creará el efecto de banda sombreada).
   d. Ajusta el color del área: clic derecho sobre el área → **Formato de serie de datos** → relleno color azul claro con 50% de transparencia.
   e. Cambia el color del área del **Lim_Inferior** a **Sin relleno** (para que solo se vea la banda entre los dos límites).

   > **Alternativa simplificada si el procedimiento anterior es complejo:** Usa el gráfico ya generado automáticamente por Excel en la hoja `Pronostico_Excel_Auto` y simplemente cópialo a la hoja `Grafico_Intervalo` usando Ctrl+C / Ctrl+V.

6. Agrega los siguientes elementos al gráfico:
   - **Título:** `Pronóstico de Demanda — Nuevo Producto (jul–dic 2024)`
   - **Eje Y:** `Unidades`
   - **Eje X:** `Período`
   - **Leyenda:** Verifica que muestre Demanda Real, Pronóstico, Intervalo de Confianza 95%
   - **Línea vertical** (opcional): Agrega una forma de línea vertical para marcar la separación entre datos históricos y pronóstico.

7. Guarda el archivo.

#### Resultado esperado

El gráfico muestra claramente tres elementos diferenciados:
- **Línea sólida azul:** demanda histórica real (ene-24 a jun-24).
- **Línea discontinua naranja o azul oscuro:** pronóstico central (jul-24 a dic-24).
- **Área sombreada celeste:** banda del intervalo de confianza al 95% entre límite inferior y superior.

Este tipo de visualización es estándar en reportes de planificación de demanda y comunica tanto la proyección central como el rango de incertidumbre operativa.

#### Verificación

✅ El gráfico tiene título y etiquetas de ejes.

✅ El área sombreada es visible y corresponde a los 6 meses de pronóstico futuro.

✅ La línea de Demanda_Real termina en jun-24 y la línea de pronóstico comienza en jun-24 o jul-24 (punto de transición visible).

---

### Paso 5 — Aplicación Manual de SES y Comparación con Pronóstico Automático

**Objetivo:** Calcular el pronóstico de los 6 meses futuros usando SES manualmente y comparar cuantitativamente con el pronóstico automático de Excel.

#### Instrucciones

1. Crea una nueva hoja llamada `Comparacion_SES_vs_Auto`.

2. Copia los 18 meses de datos históricos originales desde la hoja **Datos_Historicos** hacia las columnas A y B de esta nueva hoja.

3. En la celda **D1** escribe `alfa_SES` y en **E1** escribe `0.3` (usarás alfa = 0.30 como punto de partida; puedes ajustarlo luego).

4. Construye la columna de pronóstico SES en la columna C con los siguientes pasos:

   ```
   C1: SES_Pronostico
   
   C2: =B2   (inicialización: el primer pronóstico es igual al primer valor real)
   
   C3: =E$1*B2+(1-E$1)*C2
   (Esta fórmula aplica: F_t = α * Y_(t-1) + (1-α) * F_(t-1))
   
   Arrastra la fórmula de C3 hasta C19 (cubre todos los 18 meses históricos)
   ```

   > **Nota importante sobre referencias:** La referencia `E$1` usa `$` para fijar la fila 1, de modo que al arrastrar la fórmula hacia abajo, siempre tome el valor de alfa de E1. Verifica esto antes de arrastrar.

5. En la celda **C20**, calcula el **pronóstico para el mes 19 (jul-24)** usando la misma fórmula SES:

   ```excel
   C20: =E$1*B19+(1-E$1)*C19
   ```

6. Para los meses 20 al 25 (ago-24 a dic-24), el SES simple **no tiene tendencia ni estacionalidad**, por lo que el pronóstico futuro se mantiene constante igual al último valor calculado. En las celdas C21 a C25 escribe:

   ```excel
   C21: =C20
   C22: =C20
   C23: =C20
   C24: =C20
   C25: =C20
   ```

   > Esto refleja una limitación importante del SES: al no capturar tendencia, proyecta un valor plano hacia el futuro. Compara esto con el pronóstico automático de Excel que sí puede capturar tendencia y estacionalidad.

7. Ahora construye la tabla de comparación. En la columna E (a partir de E3) crea la siguiente estructura:

   ```
   E3: Mes_Futuro
   F3: Pronostico_SES
   G3: Pronostico_Excel_Auto
   H3: Diferencia_Absoluta
   I3: Diferencia_%
   
   E4:  jul-24
   E5:  ago-24
   E6:  sep-24
   E7:  oct-24
   E8:  nov-24
   E9:  dic-24
   ```

8. Completa las columnas F y G copiando los valores correspondientes:
   - **Columna F (SES):** Copia los valores de C20:C25.
   - **Columna G (Excel Auto):** Regresa a la hoja `Pronostico_Excel_Auto`, copia los 6 valores de pronóstico central (jul-24 a dic-24) y pégalos como **valores** en G4:G9.

9. Calcula las diferencias en H y I:

   ```excel
   H4: =ABS(F4-G4)
   I4: =H4/G4
   
   (Arrastra H4:I4 hacia abajo hasta H9:I9)
   ```

   Formatea la columna I como porcentaje con 1 decimal.

10. En la celda **E11** escribe `Diferencia promedio:` y en F11 calcula `=PROMEDIO(H4:H9)`.

11. Guarda el archivo.

#### Resultado esperado

La tabla de comparación muestra los pronósticos de ambos métodos para los 6 meses futuros. Es esperable observar que:

- El **SES proyecta un valor relativamente plano** (o ligeramente variable) porque no captura tendencia.
- El **pronóstico automático de Excel (ETS)** puede mostrar una tendencia creciente o decreciente y posiblemente un patrón estacional, si Excel lo detectó en los 18 meses de historia.
- La diferencia porcentual entre ambos métodos puede ser de 5% a 25% dependiendo de la estructura de los datos, lo cual es operativamente significativo para decisiones de inventario.

#### Verificación

✅ Las celdas C20:C25 contienen los pronósticos SES para los 6 meses futuros.

✅ La tabla de comparación E3:I9 está completa sin celdas vacías.

✅ La diferencia promedio en F11 tiene un valor calculado (no un error #DIV/0! ni #REF!).

---

### Paso 6 — Consulta a IA Generativa sobre el Método ETS

**Objetivo:** Investigar el método estadístico subyacente (ETS) que Excel usa en la Hoja de Pronóstico, documentar la respuesta de la IA y vincular los conceptos con decisiones logísticas.

#### Instrucciones

1. Abre tu herramienta de IA Generativa (Copilot en `copilot.microsoft.com`, ChatGPT en `chat.openai.com` o equivalente).

2. Escribe el siguiente **prompt** exactamente como se indica (puedes copiarlo desde aquí):

   ```
   Prompt 1 — Concepto ETS:
   
   "Soy analista de logística y estoy usando la herramienta 'Hoja de Pronóstico' 
   de Microsoft Excel. Me indica que utiliza el método ETS (Error, Trend, 
   Seasonality). Explícame en términos simples:
   1. ¿Qué significa cada componente de ETS (Error, Trend, Seasonality)?
   2. ¿Cómo difiere ETS del Suavizamiento Exponencial Simple (SES)?
   3. ¿Por qué ETS puede ser más preciso que un promedio móvil simple para 
      pronosticar demanda con estacionalidad?
   4. ¿Cuáles son las limitaciones de ETS en contextos de logística y gestión 
      de inventario?
   Responde de forma concisa, máximo 300 palabras."
   ```

3. Lee la respuesta completa de la IA. Luego escribe un **segundo prompt** de seguimiento:

   ```
   Prompt 2 — Implicaciones para inventario:
   
   "Dado que el intervalo de confianza del pronóstico ETS se amplía con el 
   horizonte temporal (el intervalo es más angosto en el mes 1 y más ancho 
   en el mes 6), ¿cómo debería un planificador de inventario usar esta 
   información para dimensionar el stock de seguridad? 
   ¿Es recomendable usar el límite superior del intervalo como demanda 
   máxima para el cálculo del stock de seguridad?"
   ```

4. Crea una nueva hoja en el libro llamada `Consulta_IA_ETS`.

5. Documenta la sesión de IA en esta hoja con la siguiente estructura:

   ```
   A1: CONSULTA A IA GENERATIVA — MÉTODO ETS
   A2: Herramienta utilizada: [nombre de la IA]
   A3: Fecha de consulta: [fecha de hoy]
   
   A5: PROMPT 1:
   B5: [Pega aquí el texto del Prompt 1]
   
   A6: RESPUESTA 1 (resumen):
   B6: [Escribe aquí un resumen de 3-5 puntos clave de la respuesta de la IA]
   
   A8: PROMPT 2:
   B8: [Pega aquí el texto del Prompt 2]
   
   A9: RESPUESTA 2 (resumen):
   B9: [Escribe aquí un resumen de 2-3 puntos clave]
   
   A11: CONCLUSIÓN OPERATIVA:
   B11: [En 2-3 oraciones: ¿Cómo cambiaría tu decisión de reabastecimiento 
         si usas el límite superior del intervalo vs. el pronóstico central?]
   ```

6. En la celda **A13** escribe la siguiente pregunta de reflexión y respóndela en B13:

   ```
   A13: PREGUNTA DE REFLEXIÓN:
   B13: "Si el CV de la demanda calculado en el Paso 1 es mayor a 0.30, 
        ¿confiarías más en el pronóstico ETS automático de Excel o en el 
        SES manual? ¿Por qué?"
   ```

7. Guarda el archivo.

#### Resultado esperado

La hoja `Consulta_IA_ETS` contiene la documentación completa de la sesión de IA con los prompts, las respuestas resumidas y la conclusión operativa. La IA debería haber explicado que ETS es una familia de modelos de suavizamiento exponencial que combina tres componentes (error, tendencia, estacionalidad), que es más sofisticado que SES porque puede capturar patrones temporales, y que el intervalo de confianza creciente refleja la incertidumbre acumulada — lo cual es directamente relevante para el dimensionamiento del stock de seguridad.

#### Verificación

✅ La hoja `Consulta_IA_ETS` existe con contenido en todas las celdas especificadas.

✅ La conclusión operativa en B11 menciona explícitamente una decisión de **inventario** o **reabastecimiento** (no es una respuesta genérica).

✅ La respuesta de reflexión en B13 hace referencia al CV calculado en el Paso 1.

---

### Paso 7 — Síntesis Final: Tabla Comparativa de Métodos

**Objetivo:** Consolidar el aprendizaje de la práctica en una tabla comparativa que contraste los tres enfoques de pronóstico trabajados.

#### Instrucciones

1. Crea una última hoja llamada `Sintesis_Comparativa`.

2. Construye la siguiente tabla comparativa (completa las celdas marcadas con `[...]` basándote en lo que observaste durante la práctica):

   | Criterio | PMS (N=3) | SES (α=0.30) | ETS Automático (Excel) |
   |---|---|---|---|
   | **Complejidad de implementación** | Baja | Media | Baja (automático) |
   | **Captura tendencia** | No | No | Sí (si la detecta) |
   | **Captura estacionalidad** | No | No | Sí (si la detecta) |
   | **Intervalo de confianza** | No genera | No genera | Sí, al 95% |
   | **Pronóstico plano a futuro** | Sí | Sí | No necesariamente |
   | **Parámetros a calibrar** | N (ventana) | α (alfa) | Automático |
   | **Pronóstico central jul-24** | [valor de prácticas anteriores] | [valor calculado en Paso 5] | [valor de hoja automática] |
   | **Recomendación de uso** | Demanda estable, sin tendencia | Demanda con cambios graduales | Demanda con tendencia o estacionalidad |

3. En las celdas debajo de la tabla, escribe una **recomendación operativa** en 3-4 oraciones:

   ```
   A15: RECOMENDACIÓN OPERATIVA PARA EL EQUIPO DE PLANIFICACIÓN:
   
   B15: [Escribe aquí tu recomendación. Debe responder: 
        ¿Cuál método usarías para este producto nuevo de 18 meses de historia? 
        ¿Cómo usarías el intervalo de confianza para definir el stock de seguridad? 
        ¿Qué información adicional necesitarías para mejorar el pronóstico?]
   ```

4. Guarda el archivo por última vez.

#### Resultado esperado

La tabla comparativa muestra claramente las diferencias entre los tres métodos en términos de capacidades, limitaciones y aplicabilidad. La recomendación operativa conecta el análisis estadístico con una decisión concreta de inventario, cerrando el ciclo de la práctica.

#### Verificación

✅ La tabla comparativa tiene los tres métodos en columnas con valores en todas las filas.

✅ La recomendación operativa menciona al menos uno de los siguientes conceptos: stock de seguridad, nivel de servicio, ciclo de reabastecimiento, o planificación de picking.

---

## 7. Validación y Pruebas Finales

Al finalizar todos los pasos, verifica que tu archivo `Lab_17_Pronostico_Completo.xlsx` cumple con los siguientes criterios de validación:

### Lista de verificación de entregables

| # | Criterio de validación | Estado |
|---|---|---|
| 1 | La hoja `Datos_Historicos` contiene los 18 meses originales sin modificaciones | ☐ |
| 2 | La hoja `Pronostico_Excel_Auto` tiene 6 filas de pronóstico futuro con 3 columnas (pronóstico + límites) | ☐ |
| 3 | La hoja `Parametros_Documentacion` tiene la tabla completa con valores en columna B y columna C | ☐ |
| 4 | La hoja `Grafico_Intervalo` contiene un gráfico con área sombreada visible para el intervalo de confianza | ☐ |
| 5 | La hoja `Comparacion_SES_vs_Auto` tiene la tabla de comparación completa con diferencias calculadas | ☐ |
| 6 | La hoja `Consulta_IA_ETS` tiene los prompts, respuestas resumidas y conclusión operativa | ☐ |
| 7 | La hoja `Sintesis_Comparativa` tiene la tabla de tres métodos y la recomendación operativa | ☐ |
| 8 | El archivo está guardado como `Lab_17_Pronostico_Completo.xlsx` | ☐ |

### Prueba de coherencia numérica

Ejecuta esta verificación rápida para confirmar que los cálculos son coherentes:

1. Abre la hoja `Comparacion_SES_vs_Auto`.
2. Verifica que el **pronóstico SES para jul-24** (celda C20) es **diferente** del promedio simple de los 18 meses (el SES pondera más los datos recientes, por lo que debe diferir del promedio aritmético).
3. Verifica que el **pronóstico central de Excel** para jul-24 es **diferente** del pronóstico SES para el mismo mes (si son idénticos, puede indicar que Excel también usó SES sin tendencia, lo cual es posible si los datos no muestran tendencia clara — en ese caso, documenta esta observación en la hoja `Parametros_Documentacion`).
4. Confirma que la amplitud del intervalo de confianza en dic-24 es **mayor** que en jul-24.

---

## 8. Solución de Problemas

### Problema 1: El botón "Hoja de Pronóstico" no aparece en la pestaña Datos

**Síntoma:** El grupo **Pronóstico** no es visible en la pestaña **Datos**, o el botón aparece atenuado (en gris) y no es clickeable.

**Causa probable:** La versión de Excel instalada es anterior a 2016, o el archivo de datos no está en formato de tabla/rango reconocible, o Excel no detecta una columna de fechas válida en el rango seleccionado.

**Solución paso a paso:**

1. **Verificar versión de Excel:** Ve a **Archivo → Cuenta → Acerca de Excel**. Confirma que la versión es 2016 o superior. Si es anterior, esta herramienta no está disponible — usa el método alternativo de la Herramienta de Análisis de Datos (Datos → Análisis de datos → Promedio móvil) y documenta que la Hoja de Pronóstico no está disponible en tu versión.

2. **Verificar formato de fechas:** Selecciona la columna A (Fecha). Ve a **Inicio → Número** y confirma que el formato es **Fecha** (no Texto). Si las fechas están como texto, selecciona la columna, ve a **Datos → Texto en columnas → Finalizar** para convertirlas.

3. **Verificar que el cursor está dentro del rango:** Haz clic en cualquier celda dentro del rango A1:B19 antes de abrir la herramienta. Si el cursor está fuera del rango de datos, Excel no puede inferir qué serie pronosticar.

4. **Si el problema persiste:** Solicita al instructor el archivo de checkpoint `Lab_17_CP1_HojaPronostico.xlsx` que contiene la hoja `Pronostico_Excel_Auto` ya generada, para continuar con los Pasos 3 en adelante.

---

### Problema 2: El gráfico con área sombreada no muestra la banda del intervalo de confianza correctamente

**Síntoma:** El gráfico muestra tres líneas separadas (Pronóstico, Límite Inferior, Límite Superior) en lugar de una línea central con área sombreada. O bien, el área sombreada cubre toda la zona desde el eje X hasta el límite superior, en lugar de solo la banda entre límite inferior y superior.

**Causa probable:** El tipo de gráfico para las series de límites no fue configurado correctamente como **Área apilada** (stacked area), o el orden de las series en el gráfico es incorrecto.

**Solución paso a paso:**

1. **Verificar el orden de las series:** Haz clic derecho en el gráfico → **Seleccionar datos**. El orden correcto de las series debe ser: (1) Lim_Inferior, (2) Lim_Superior, (3) Pronóstico, (4) Demanda_Real. Usa los botones **Subir/Bajar** para reordenar si es necesario.

2. **Configurar Lim_Inferior como Área sin relleno:** Haz clic derecho sobre la serie Lim_Inferior → **Formato de serie de datos** → **Relleno** → selecciona **Sin relleno**. Esto hace que el área inferior sea transparente.

3. **Configurar Lim_Superior como Área apilada con color:** Haz clic derecho sobre la serie Lim_Superior → **Cambiar tipo de gráfico de serie** → selecciona **Área apilada**. Luego en **Formato de serie de datos** → **Relleno** → **Relleno sólido** → elige azul claro → ajusta **Transparencia** al 50%.

4. **Alternativa rápida (si los pasos anteriores consumen demasiado tiempo):** Usa directamente el gráfico generado automáticamente por Excel en la hoja `Pronostico_Excel_Auto`. Haz clic en el gráfico → Ctrl+C → ve a la hoja `Grafico_Intervalo` → Ctrl+V. Luego ajusta el título y las etiquetas. Este gráfico ya tiene el área sombreada correctamente configurada por Excel.

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones de cierre:

1. **Guardar el archivo final:** Presiona **Ctrl + S** para guardar `Lab_17_Pronostico_Completo.xlsx` por última vez. Confirma que el archivo tiene **7 hojas** (Datos_Historicos, Pronostico_Excel_Auto, Parametros_Documentacion, Grafico_Intervalo, Comparacion_SES_vs_Auto, Consulta_IA_ETS, Sintesis_Comparativa).

2. **Crear copia de respaldo:** Guarda una copia adicional con el nombre `Lab_17_Pronostico_Completo_BACKUP.xlsx` en una carpeta diferente o en la nube (OneDrive, Google Drive).

3. **Cerrar la sesión de IA Generativa:** Si usaste una herramienta de IA en el navegador, cierra la pestaña o la sesión. No es necesario guardar el historial de chat, ya que la información relevante fue documentada en la hoja `Consulta_IA_ETS`.

4. **No eliminar ningún archivo intermedio:** El archivo `Lab_17_Demanda_NuevoProducto.xlsx` original debe mantenerse intacto. Las Prácticas 18 y siguientes pueden requerir los datos originales como referencia.

5. **Verificación final de estructura:** Abre el panel de hojas (pestañas inferiores) y confirma visualmente que las 7 hojas existen con los nombres correctos.

---

## 10. Resumen y Recursos Adicionales

### Síntesis de lo aprendido

En esta práctica aplicaste tres niveles de sofisticación en el pronóstico de demanda sobre el mismo dataset:

1. **Nivel básico — SES manual:** Construiste el pronóstico aplicando manualmente la fórmula de suavizamiento exponencial, controlando el parámetro alfa y entendiendo cada paso del cálculo. La limitación principal: proyección plana sin tendencia ni estacionalidad.

2. **Nivel intermedio — Hoja de Pronóstico de Excel (ETS):** Usaste la herramienta automática que aplica internamente el método ETS, capturando tendencia y estacionalidad cuando están presentes en los datos, y generando intervalos de confianza al 95%. La ventaja clave: automatización y mayor precisión estadística. La limitación: menor control sobre los parámetros y riesgo de sobreajuste en series cortas.

3. **Nivel conceptual — Investigación con IA Generativa:** Usaste la IA para comprender el fundamento teórico del método ETS y vincularlo con decisiones operativas de inventario, específicamente el uso del intervalo de confianza para dimensionar el stock de seguridad.

### Conexión con decisiones logísticas

La pregunta central de toda práctica de forecasting logístico es: **¿cómo impacta este análisis en una decisión operativa?** En esta práctica, las respuestas concretas son:

- El **pronóstico central** (punto medio del intervalo) se usa como base para el **plan de reabastecimiento**.
- El **límite superior del intervalo al 95%** puede usarse para dimensionar el **stock de seguridad** ante variabilidad de demanda.
- La **amplitud creciente del intervalo** en horizontes largos indica que los pedidos a 6 meses plazo requieren mayor colchón de seguridad que los pedidos a 1 mes.
- La **comparación entre SES y ETS** permite al planificador elegir el modelo más adecuado según la estructura de la demanda del producto.

### Recursos adicionales recomendados

| Recurso | Descripción | Acceso |
|---|---|---|
| Hyndman & Athanasopoulos — *Forecasting: Principles and Practice* | Capítulo 8: Métodos ETS — explicación completa y gratuita | [otexts.com/fpp3/](https://otexts.com/fpp3/) |
| Microsoft Support — Hoja de Pronóstico de Excel | Documentación oficial con ejemplos paso a paso | [support.microsoft.com](https://support.microsoft.com/es-es/office/crear-una-hoja-de-previsi%C3%B3n-en-excel-para-windows-22c500da-6da7-45e5-bfdc-60a7062329fd) |
| Chopra & Meindl — *Supply Chain Management* | Capítulo 7: Pronóstico de la demanda en cadenas de suministro | Biblioteca institucional o Pearson |
| Canal YouTube — ExcelTotal | Tutoriales en español sobre Hoja de Pronóstico y gráficos avanzados | YouTube: ExcelTotal |

### Preparación para la Práctica 18

La siguiente práctica utilizará el archivo `Lab_17_Pronostico_Completo.xlsx` como punto de partida, específicamente los pronósticos generados en las hojas `Comparacion_SES_vs_Auto` y `Pronostico_Excel_Auto`. Asegúrate de tener el archivo guardado y accesible antes de la próxima sesión.

---

> **Recordatorio del instructor:** Verificar que todos los participantes tengan el archivo `Lab_17_Pronostico_Completo.xlsx` guardado correctamente con las 7 hojas antes de cerrar la sesión. Distribuir el archivo de checkpoint `Lab_17_CP1_HojaPronostico.xlsx` a quienes no pudieron completar el Paso 2 por problemas de versión de Excel.

---
LAB_END---

---

---LAB_START---
LAB_ID: 03-00-06
---MARKDOWN---
# Práctica 18 — Estimación de Volumen de Líneas de Picking a partir del Pronóstico de Demanda

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 20 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Aplicar (Apply) |
| **Práctica número** | 18 |
| **Módulo** | 3 — Métodos Cuantitativos de Pronóstico |
| **Lección de referencia** | 3.1 — Método de Promedio Móvil |
| **Herramientas principales** | Microsoft Excel 2016 o superior |
| **Archivos de entrada** | `Lab18_Pronostico_Demanda.xlsx` (proporcionado por el instructor) |
| **Archivo de salida** | `Lab18_Picking_Capacidad_[TuNombre].xlsx` |

---

## 2. Descripción General

En esta práctica traducirás el pronóstico de demanda mensual —generado con el método de **Promedio Móvil Simple** aprendido en la Lección 3.1— en una estimación operativa del **volumen de líneas de picking** requeridas en el almacén durante los próximos 6 meses. Integrarás parámetros logísticos reales (unidades por línea, porcentaje de devoluciones y factor de re-picking) para construir un modelo de planeación de capacidad completamente dinámico en Excel. El resultado final te permitirá identificar meses con riesgo de **sobrecarga operativa** y formular recomendaciones concretas para el equipo de almacén, cerrando el ciclo entre el pronóstico estadístico y la decisión operativa logística.

---

## 3. Objetivos de Aprendizaje

Al finalizar esta práctica serás capaz de:

- [ ] **Aplicar** la fórmula de estimación de líneas de picking integrando parámetros logísticos (unidades por línea, devoluciones y re-picking) sobre un pronóstico de demanda mensual.
- [ ] **Construir** en Excel un modelo parametrizado con referencias absolutas que vincule dinámicamente el pronóstico de demanda con los requerimientos operativos de picking.
- [ ] **Visualizar** la contribución por producto al volumen total mensual de picking mediante un gráfico de barras apiladas.
- [ ] **Identificar** meses con riesgo de sobrecarga operativa comparando el volumen estimado contra la capacidad máxima de picking definida, utilizando formato condicional como alerta.
- [ ] **Formular** al menos dos recomendaciones operativas justificadas a partir del análisis de capacidad.

---

## 4. Prerrequisitos

### Conocimiento previo requerido
- Comprensión del proceso de picking en almacenes (qué es una línea de picking, qué implica una devolución y un re-picking).
- Haber completado al menos una de las Prácticas 13 a 17 (construcción de pronóstico con Promedio Móvil u otro método cuantitativo).
- Manejo de referencias absolutas (`$`) en Excel para fijar parámetros en fórmulas.
- Familiaridad con la creación de gráficos de barras y aplicación de formato condicional en Excel.

### Acceso y archivos requeridos
- Microsoft Excel 2016 o superior instalado y funcional.
- Archivo `Lab18_Pronostico_Demanda.xlsx` descargado y disponible en tu escritorio o carpeta de trabajo.
- Si no completaste las prácticas anteriores, solicita al instructor el archivo de *checkpoint* `Lab18_Checkpoint_Inicio.xlsx`.

> **⚠️ Nota de secuencialidad:** Esta práctica utiliza los pronósticos generados en las Prácticas 13–17. Si tu pronóstico personal difiere de los valores de referencia, usa los datos del archivo proporcionado por el instructor para garantizar resultados comparables con el grupo.

---

## 5. Entorno de Laboratorio

### Hardware recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Resolución de pantalla | 1366 × 768 | 1920 × 1080 |
| Almacenamiento libre | 2 GB | 2 GB |
| Conexión a internet | Opcional para esta práctica | 10 Mbps (para IA Generativa en prácticas siguientes) |

### Software requerido

| Software | Versión | Notas |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | Necesario para funciones `PROMEDIO`, `SUMA`, formato condicional y gráficos |
| Archivo de datos | `Lab18_Pronostico_Demanda.xlsx` | Proporcionado por el instructor |

### Configuración inicial (antes de comenzar)

1. Abre el Explorador de archivos y navega a la carpeta donde guardaste `Lab18_Pronostico_Demanda.xlsx`.
2. Haz doble clic para abrirlo en Excel.
3. Guarda inmediatamente una copia con tu nombre:
   - **Archivo → Guardar como → `Lab18_Picking_Capacidad_[TuNombre].xlsx`**
4. Verifica que el archivo contiene las siguientes hojas:
   - `Pronóstico_Demanda` — Pronóstico mensual en unidades para 5 productos × 6 meses.
   - `Parámetros_Logísticos` — Tabla de parámetros por producto.
   - `Picking_Model` — Hoja en blanco donde construirás el modelo (puede estar parcialmente estructurada).
5. Familiarízate con los datos durante 60 segundos antes de iniciar el Paso 1.

### Datos de referencia del archivo

**Tabla A — Pronóstico de demanda mensual (unidades)**

| SKU / Producto | Mes 1 | Mes 2 | Mes 3 | Mes 4 | Mes 5 | Mes 6 |
|---|---|---|---|---|---|---|
| SKU-01 Aceite Vegetal 1L | 1 820 | 1 890 | 1 950 | 2 010 | 1 980 | 2 050 |
| SKU-02 Arroz Blanco 5kg | 3 400 | 3 520 | 3 480 | 3 600 | 3 650 | 3 720 |
| SKU-03 Leche UHT 1L | 2 150 | 2 200 | 2 180 | 2 250 | 2 300 | 2 270 |
| SKU-04 Azúcar Blanca 1kg | 1 600 | 1 650 | 1 630 | 1 700 | 1 680 | 1 750 |
| SKU-05 Harina de Trigo 1kg | 980 | 1 010 | 1 040 | 1 020 | 1 060 | 1 080 |

**Tabla B — Parámetros logísticos por producto**

| SKU / Producto | Unidades por Línea | % Devoluciones | Factor Re-picking |
|---|---|---|---|
| SKU-01 Aceite Vegetal 1L | 12 | 3.5% | 1.02 |
| SKU-02 Arroz Blanco 5kg | 20 | 2.0% | 1.01 |
| SKU-03 Leche UHT 1L | 24 | 4.0% | 1.03 |
| SKU-04 Azúcar Blanca 1kg | 15 | 2.5% | 1.02 |
| SKU-05 Harina de Trigo 1kg | 18 | 3.0% | 1.01 |

**Capacidad máxima de picking del almacén:** **1 200 líneas/mes** (dato fijo para análisis de capacidad).

> **Nota pedagógica:** Los valores de pronóstico de la Tabla A fueron generados aplicando el **Promedio Móvil Simple con N=3** sobre 24 meses de historial, exactamente como se practicó en las sesiones anteriores. Observa cómo el pronóstico "suavizado" del PM-3 produce valores relativamente estables mes a mes, lo cual es consistente con lo aprendido en la Lección 3.1.

---

## 6. Instrucciones Paso a Paso

---

### Paso 1 — Revisar y organizar los datos de entrada en la hoja `Picking_Model`

**Objetivo:** Estructurar correctamente la hoja de trabajo antes de construir las fórmulas, garantizando que los parámetros logísticos estén referenciados de forma absoluta.

**Instrucciones:**

1. Haz clic en la hoja **`Picking_Model`** en la parte inferior de Excel.

2. En la celda **A1**, escribe el título:
   ```
   MODELO DE ESTIMACIÓN DE LÍNEAS DE PICKING — HORIZONTE 6 MESES
   ```
   Aplica **negrita** y **tamaño de fuente 14**. Fusiona las celdas A1:K1 (selecciona el rango → pestaña Inicio → Combinar y centrar).

3. Construye el encabezado de la tabla principal comenzando en la fila **3**:

   | Celda | Contenido |
   |---|---|
   | A3 | `SKU` |
   | B3 | `Producto` |
   | C3 | `Und/Línea` |
   | D3 | `% Dev` |
   | E3 | `F.Repick` |
   | F3 | `Mes 1` |
   | G3 | `Mes 2` |
   | H3 | `Mes 3` |
   | I3 | `Mes 4` |
   | J3 | `Mes 5` |
   | K3 | `Mes 6` |

4. Aplica **fondo azul oscuro** y **texto blanco** a toda la fila 3 (selecciona A3:K3 → Color de relleno → Azul oscuro; Color de fuente → Blanco).

5. Ingresa los datos de los 5 productos en las filas **4 a 8**, copiando la información de la hoja `Parámetros_Logísticos`:
   - Columna A: SKU-01 a SKU-05
   - Columna B: Nombre del producto
   - Columna C: Unidades por línea (12, 20, 24, 15, 18)
   - Columna D: % Devoluciones como decimal (0.035, 0.020, 0.040, 0.025, 0.030)
   - Columna E: Factor Re-picking (1.02, 1.01, 1.03, 1.02, 1.01)

   > **Importante:** Formatea las celdas D4:D8 como **Porcentaje con 1 decimal** (selecciona → Ctrl+1 → Porcentaje → 1 decimal). Esto hará que los valores se muestren como 3.5%, 2.0%, etc., aunque internamente sean decimales.

6. En la fila **9**, escribe en la columna A: `TOTAL LÍNEAS/MES` y aplica negrita.

7. Reserva las filas **11 a 14** para la sección de análisis de capacidad (las completarás en el Paso 4). Por ahora escribe en A11: `ANÁLISIS DE CAPACIDAD` con negrita.

8. Guarda el archivo: **Ctrl + S**.

**Resultado esperado:** Una hoja estructurada con encabezados claros, datos de los 5 productos en filas 4–8 con parámetros logísticos visibles, y espacio reservado para las secciones posteriores.

**Verificación:** Confirma que las columnas C, D y E contienen los parámetros correctos para los 5 productos antes de continuar. Compara con la Tabla B de la sección de datos de referencia.

---

### Paso 2 — Construir la fórmula de estimación de líneas de picking

**Objetivo:** Implementar la fórmula central del modelo con referencias absolutas para los parámetros, de modo que sea dinámica y reutilizable para todos los productos y meses.

**Instrucciones:**

1. Antes de escribir la fórmula, comprende su estructura lógica:

   ```
   Líneas_Picking = (Demanda_Pronosticada / Unidades_por_Línea) × (1 + % Devoluciones) × Factor_Repicking
   ```

   - **`Demanda_Pronosticada / Unidades_por_Línea`**: Convierte unidades en líneas base (cuántas líneas se necesitan para despachar esa demanda).
   - **`× (1 + % Devoluciones)`**: Agrega líneas adicionales por devoluciones que generan re-trabajo.
   - **`× Factor_Repicking`**: Multiplica por el factor de re-picking que captura re-procesos operativos adicionales (etiquetado erróneo, picking fallido, etc.).

2. Haz clic en la celda **F4** (primera celda de datos: SKU-01, Mes 1).

3. Escribe la siguiente fórmula. Presta atención a las referencias absolutas (`$`) en las columnas de parámetros:

   ```excel
   =(Pronóstico_Demanda!B2/$C4)*(1+$D4)*$E4
   ```

   > **Explicación de referencias:**
   > - `Pronóstico_Demanda!B2` → Valor de demanda pronosticada del Mes 1 para SKU-01 (referencia a la otra hoja; ajusta la celda exacta según la estructura de tu archivo).
   > - `$C4` → Referencia mixta: la columna C (Und/Línea) está fija con `$`, pero la fila 4 puede cambiar al arrastrar hacia abajo.
   > - `$D4` → Igual: columna D fija (% Devoluciones), fila variable.
   > - `$E4` → Igual: columna E fija (Factor Re-picking), fila variable.

   > **Alternativa si los datos están en la misma hoja:** Si copiaste los valores de pronóstico directamente en la hoja `Picking_Model` (por ejemplo en un rango separado), ajusta la referencia de demanda según corresponda. Por ejemplo, si el pronóstico del SKU-01 Mes 1 está en la celda N4, la fórmula sería:
   > ```excel
   > =($N4/$C4)*(1+$D4)*$E4
   > ```

4. Presiona **Enter** y verifica que el resultado para SKU-01, Mes 1 sea aproximadamente **161.8 líneas** (cálculo de referencia: 1820/12 × 1.035 × 1.02 ≈ 161.8).

5. Regresa a la celda F4. Ahora arrastra la fórmula hacia la **derecha** hasta la celda **K4** (cubriendo los 6 meses para SKU-01). Verifica que las referencias de demanda cambien de columna correctamente (Mes 1 → Mes 2 → ... → Mes 6) mientras los parámetros `$C4`, `$D4`, `$E4` permanecen fijos en sus columnas.

6. Selecciona el rango **F4:K4** y arrastra hacia **abajo** hasta la fila **8** (cubriendo los 5 productos). Verifica que las referencias de parámetros cambien de fila (de `$C4` a `$C5`, `$C6`, etc.) mientras las columnas permanecen fijas.

7. Formatea el rango **F4:K8** como **Número con 1 decimal** (Ctrl+1 → Número → 1 decimal). Esto facilita la lectura de las líneas de picking.

8. Aplica **fondo amarillo claro** al rango F4:K8 para distinguirlo visualmente como zona de resultados calculados.

**Tabla de verificación de resultados esperados (líneas de picking por producto y mes):**

| SKU | Mes 1 | Mes 2 | Mes 3 | Mes 4 | Mes 5 | Mes 6 |
|---|---|---|---|---|---|---|
| SKU-01 | 161.8 | 167.9 | 173.2 | 178.6 | 175.9 | 182.1 |
| SKU-02 | 174.5 | 180.6 | 178.5 | 184.7 | 187.3 | 190.8 |
| SKU-03 | 96.3 | 98.6 | 97.7 | 100.8 | 103.1 | 101.7 |
| SKU-04 | 109.2 | 112.6 | 111.2 | 116.1 | 114.7 | 119.5 |
| SKU-05 | 55.5 | 57.3 | 59.0 | 57.8 | 60.1 | 61.2 |

> **Tolerancia:** Una diferencia de ±0.5 líneas respecto a los valores de referencia es aceptable y puede deberse a redondeo en los datos de pronóstico de entrada.

**Resultado esperado:** Tabla completa de 5 productos × 6 meses con líneas de picking estimadas, con fórmulas dinámicas que respetan los parámetros logísticos por producto.

**Verificación:** Modifica temporalmente el valor de `Und/Línea` del SKU-01 (celda C4) de 12 a 10. Todos los valores de la fila 4 deben aumentar proporcionalmente. Restaura el valor a 12 antes de continuar.

---

### Paso 3 — Calcular el total de líneas de picking por mes y construir el gráfico de barras apiladas

**Objetivo:** Consolidar el volumen total mensual de picking y visualizar la contribución de cada producto mediante un gráfico de barras apiladas.

**Instrucciones:**

1. Haz clic en la celda **F9** (fila de TOTAL LÍNEAS/MES, Mes 1).

2. Escribe la fórmula de suma:
   ```excel
   =SUMA(F4:F8)
   ```

3. Arrastra la fórmula desde F9 hacia la derecha hasta **K9** para cubrir los 6 meses.

4. Formatea el rango **F9:K9** con:
   - **Negrita**
   - **Fondo azul oscuro** y **texto blanco**
   - **Número con 1 decimal**

5. Verifica que los totales sean aproximadamente los siguientes:

   | Mes 1 | Mes 2 | Mes 3 | Mes 4 | Mes 5 | Mes 6 |
   |---|---|---|---|---|---|
   | 597.3 | 617.0 | 619.6 | 638.0 | 641.1 | 655.3 |

6. Ahora construirás el **gráfico de barras apiladas**. Selecciona el rango de datos para el gráfico:
   - Mantén presionada la tecla **Ctrl** y selecciona simultáneamente:
     - **B3:B8** (nombres de productos, incluyendo encabezado)
     - **F3:K8** (datos de líneas de picking por mes, incluyendo encabezados de mes)

   > **Tip:** Selecciona primero B3:B8, luego mantén Ctrl presionado y selecciona F3:K8.

7. Ve a la pestaña **Insertar** → grupo **Gráficos** → haz clic en **Insertar gráfico de columnas o de barras** → selecciona **Columna apilada** (la segunda opción en la primera fila de columnas).

8. Excel generará el gráfico. Ahora personalízalo:

   a. **Título del gráfico:** Haz doble clic sobre el título y escribe:
   ```
   Volumen de Líneas de Picking por Producto y Mes
   ```

   b. **Etiquetas de eje horizontal:** Verifica que muestren "Mes 1", "Mes 2", ..., "Mes 6". Si no es así, haz clic derecho sobre el gráfico → **Seleccionar datos** → edita las etiquetas del eje horizontal para apuntar al rango F3:K3.

   c. **Leyenda:** Confirma que la leyenda muestre los 5 productos. Si aparecen "Serie 1", "Serie 2", etc., haz clic derecho → **Seleccionar datos** → edita cada serie para que el nombre apunte a las celdas B4:B8 respectivamente.

   d. **Colores:** Asigna un color diferente a cada producto haciendo clic sobre cada barra del gráfico. Se recomienda usar colores contrastantes (azul, naranja, verde, rojo, morado).

   e. **Etiquetas de datos (opcional pero recomendado):** Haz clic derecho sobre cualquier segmento de barra → **Agregar etiquetas de datos**. Esto facilita la lectura del volumen por producto.

9. Mueve el gráfico debajo de la tabla principal (aproximadamente desde la fila 16 hasta la fila 32) para que no tape los datos.

10. Guarda el archivo: **Ctrl + S**.

**Resultado esperado:** Un gráfico de barras apiladas que muestra claramente cómo cada uno de los 5 productos contribuye al volumen total de picking en cada mes, con una tendencia visual de crecimiento progresivo hacia el Mes 6.

**Verificación:** Observa el gráfico y responde mentalmente: ¿Qué producto genera más líneas de picking consistentemente? (Respuesta esperada: SKU-02 Arroz Blanco, por su alto volumen de demanda relativo a sus unidades por línea.) ¿Se observa algún mes con un salto notable? Esta reflexión te prepara para el análisis de capacidad del siguiente paso.

---

### Paso 4 — Construir la sección de análisis de capacidad con formato condicional

**Objetivo:** Comparar el volumen estimado de picking contra la capacidad máxima del almacén e identificar meses con riesgo de sobrecarga operativa, utilizando formato condicional como sistema de alerta visual.

**Instrucciones:**

1. Desplázate a la fila **11** de la hoja `Picking_Model` donde reservaste el espacio para el análisis de capacidad.

2. Construye la siguiente estructura en las filas 11 a 15:

   | Celda | Contenido | Formato |
   |---|---|---|
   | A11 | `ANÁLISIS DE CAPACIDAD DE PICKING` | Negrita, tamaño 12, fondo gris claro |
   | A12 | `Capacidad Máxima (líneas/mes)` | Normal |
   | A13 | `Total Líneas Estimadas` | Normal |
   | A14 | `Utilización de Capacidad (%)` | Normal |
   | A15 | `Estado` | Negrita |

3. En la celda **B12**, ingresa el valor de capacidad máxima:
   ```
   1200
   ```
   Aplica **negrita** y **fondo verde claro** a esta celda. Este es el parámetro clave del análisis.

4. En la celda **F12**, ingresa la misma capacidad máxima con referencia absoluta para que sea reutilizable en los 6 meses:
   ```excel
   =$B$12
   ```
   Arrastra esta fórmula desde F12 hasta **K12**.

5. En la celda **F13**, vincula el total de líneas calculado en la fila 9:
   ```excel
   =F9
   ```
   Arrastra desde F13 hasta **K13**.

6. En la celda **F14**, calcula el **porcentaje de utilización de capacidad**:
   ```excel
   =F13/$B$12
   ```
   Arrastra desde F14 hasta **K14**. Formatea el rango F14:K14 como **Porcentaje con 1 decimal**.

7. En la celda **F15**, construye una fórmula de estado que clasifique automáticamente cada mes:
   ```excel
   =SI(F13>$B$12,"⚠ SOBRECARGA",SI(F13>$B$12*0.85,"⚡ ALERTA",SI(F13>$B$12*0.70,"✅ NORMAL","🟢 HOLGADO")))
   ```
   Arrastra desde F15 hasta **K15**.

   > **Lógica de los umbrales:**
   > - **> 100% capacidad** → SOBRECARGA (riesgo operativo crítico)
   > - **85% – 100%** → ALERTA (zona de precaución, considerar refuerzo)
   > - **70% – 85%** → NORMAL (operación eficiente)
   > - **< 70%** → HOLGADO (capacidad subutilizada)

   > **Nota de compatibilidad:** Si tu versión de Excel no soporta emojis en fórmulas, reemplaza los emojis por texto simple: `"SOBRECARGA"`, `"ALERTA"`, `"NORMAL"`, `"HOLGADO"`.

8. Ahora aplica **formato condicional** al rango **F15:K15** para colorear automáticamente cada celda según el estado:

   a. Selecciona F15:K15.
   b. Ve a **Inicio → Formato condicional → Nueva regla**.
   c. Selecciona **"Aplicar formato únicamente a las celdas que contengan"**.
   d. Configura la primera regla: Texto que contiene → `SOBRECARGA` → Formato: fondo rojo, texto blanco, negrita. Clic en **Aceptar**.
   e. Repite para: `ALERTA` → fondo naranja, texto negro, negrita.
   f. Repite para: `NORMAL` → fondo verde claro, texto negro.
   g. Repite para: `HOLGADO` → fondo azul claro, texto negro.

9. **Adicionalmente**, aplica formato condicional al rango **F14:K14** (porcentaje de utilización):

   a. Selecciona F14:K14.
   b. Ve a **Inicio → Formato condicional → Escalas de color**.
   c. Selecciona la escala **Verde – Amarillo – Rojo** (verde para valores bajos, rojo para valores altos). Esto crea un mapa de calor visual del nivel de utilización.

10. Aplica **formato condicional de barras de datos** al rango **F13:K13** (total de líneas):

    a. Selecciona F13:K13.
    b. Ve a **Inicio → Formato condicional → Barras de datos → Barra de datos azul sólida**.

11. Guarda el archivo: **Ctrl + S**.

**Tabla de verificación del análisis de capacidad:**

| Indicador | Mes 1 | Mes 2 | Mes 3 | Mes 4 | Mes 5 | Mes 6 |
|---|---|---|---|---|---|---|
| Capacidad máxima | 1 200 | 1 200 | 1 200 | 1 200 | 1 200 | 1 200 |
| Total líneas estimadas | ~597 | ~617 | ~620 | ~638 | ~641 | ~655 |
| Utilización (%) | ~49.8% | ~51.4% | ~51.6% | ~53.2% | ~53.4% | ~54.6% |
| Estado esperado | HOLGADO | HOLGADO | HOLGADO | HOLGADO | HOLGADO | HOLGADO |

> **Reflexión importante:** Con la capacidad actual de 1 200 líneas/mes, el almacén opera con ~50% de utilización. Esto puede parecer positivo, pero plantea preguntas operativas: ¿Es realista esta capacidad? ¿Hay estacionalidad no capturada? ¿Qué pasaría si se incorporan los 3 SKUs restantes del dataset completo (8 SKUs)? Esta reflexión es la base de las recomendaciones del Paso 5.

**Resultado esperado:** Sección de análisis de capacidad completa con colores de alerta, porcentajes de utilización y estados automáticos que cambian dinámicamente si se modifican los parámetros de entrada.

**Verificación:** Cambia temporalmente el valor de capacidad máxima en B12 de 1200 a 650. Observa cómo los estados cambian a "ALERTA" o "SOBRECARGA" para los meses con mayor volumen. Restaura el valor a 1200 antes de continuar.

---

### Paso 5 — Redactar las recomendaciones operativas

**Objetivo:** Cerrar el ciclo analítico formulando recomendaciones operativas concretas y justificadas a partir del modelo construido, vinculando el pronóstico estadístico con decisiones de gestión de almacén.

**Instrucciones:**

1. Desplázate a la fila **34** (o debajo del gráfico) y crea una nueva sección de recomendaciones.

2. En la celda **A34**, escribe:
   ```
   RECOMENDACIONES OPERATIVAS BASADAS EN EL ANÁLISIS DE PICKING
   ```
   Aplica negrita, tamaño 13, fondo azul oscuro, texto blanco. Fusiona A34:K34.

3. A partir de la fila **36**, escribe al menos **4 recomendaciones operativas** en la columna A. Cada recomendación debe:
   - Estar numerada.
   - Mencionar un hallazgo específico del análisis (producto, mes, porcentaje, líneas).
   - Proponer una acción operativa concreta.
   - Ser redactada en 2–3 líneas máximo.

   **Recomendaciones de referencia (puedes adaptarlas o enriquecerlas con tus propias observaciones):**

   ---

   **Recomendación 1 — Capacidad actual con margen operativo suficiente para el horizonte de 6 meses:**
   > El volumen total de líneas de picking estimado para los próximos 6 meses oscila entre 597 y 655 líneas/mes, lo que representa entre el 49.8% y el 54.6% de la capacidad máxima declarada (1 200 líneas/mes). Se recomienda **revisar si la capacidad declarada es realista** o si incluye turnos adicionales, ya que una utilización tan baja podría indicar subutilización de recursos o una capacidad sobredimensionada. Validar con el jefe de almacén.

   **Recomendación 2 — SKU-02 (Arroz Blanco 5kg) como producto de mayor impacto en picking:**
   > El SKU-02 genera consistentemente el mayor volumen de líneas de picking (entre 174 y 191 líneas/mes), representando aproximadamente el **29%** del total mensual. Se recomienda **priorizar la ubicación de este SKU en zonas de picking de alta accesibilidad** (posiciones de flujo rápido, nivel ergonómico) para reducir el tiempo de ciclo de picking y minimizar errores operativos.

   **Recomendación 3 — Tendencia creciente hacia el Mes 6 requiere planificación anticipada:**
   > El volumen total de líneas de picking muestra una tendencia de crecimiento progresivo: de 597 líneas en el Mes 1 a 655 en el Mes 6 (+9.7%). Aunque no representa sobrecarga, se recomienda **planificar con anticipación la asignación de personal de picking para los Meses 5 y 6**, considerando la incorporación de los 3 SKUs adicionales no incluidos en este análisis, que podrían elevar el volumen total por encima del umbral de alerta (1 020 líneas = 85% de capacidad).

   **Recomendación 4 — Integrar factor de estacionalidad en el modelo para mayor precisión:**
   > El pronóstico utilizado como base fue generado con Promedio Móvil Simple (N=3), que suaviza la demanda pero **no captura picos estacionales**. Si alguno de los 5 productos tiene comportamiento estacional (por ejemplo, mayor demanda en temporada escolar o navideña), las líneas de picking reales podrían superar significativamente las estimadas. Se recomienda revisar el historial de 24 meses para detectar patrones estacionales antes de comprometer la capacidad operativa del almacén.

   ---

4. Formatea las recomendaciones con:
   - Números en negrita (1., 2., 3., 4.)
   - Texto del cuerpo en fuente normal, tamaño 11
   - Fondo alternado: filas pares con gris muy claro, filas impares sin relleno

5. Añade en la fila **42** (o al final) una nota metodológica:

   ```
   NOTA METODOLÓGICA: Las estimaciones de líneas de picking se calcularon con la fórmula:
   Líneas_Picking = (Demanda_Pronosticada / Unidades_por_Línea) × (1 + % Devoluciones) × Factor_Repicking
   El pronóstico de demanda fue generado con Promedio Móvil Simple N=3 sobre 24 meses de historial.
   Capacidad máxima de referencia: 1,200 líneas/mes.
   ```

6. Guarda el archivo: **Ctrl + S**.

**Resultado esperado:** Sección de recomendaciones operativas con al menos 4 recomendaciones concretas, vinculadas a datos específicos del análisis, redactadas en lenguaje operativo logístico (no estadístico abstracto).

**Verificación:** Revisa que cada recomendación mencione al menos un número o porcentaje específico del análisis. Las recomendaciones genéricas sin respaldo cuantitativo no son aceptables en este contexto profesional.

---

## 7. Validación y Pruebas del Modelo

Una vez completados todos los pasos, realiza las siguientes pruebas de validación para confirmar que el modelo funciona correctamente:

### Prueba 1 — Sensibilidad de parámetros (2 minutos)

1. Cambia el **% de Devoluciones** del SKU-03 (celda D6) de 4.0% a 8.0%.
2. Verifica que **todos los valores de líneas de picking del SKU-03** (fila 6, columnas F a K) aumenten proporcionalmente.
3. Confirma que el **total de líneas por mes** (fila 9) también se actualice automáticamente.
4. Verifica que el **porcentaje de utilización** (fila 14) cambie en consecuencia.
5. Restaura el valor original: D6 = 0.04 (4.0%).

**Resultado esperado:** Todos los valores dependientes se actualizan en cascada sin necesidad de intervención manual. Esto confirma que las referencias absolutas y relativas están correctamente configuradas.

### Prueba 2 — Cambio de capacidad máxima (1 minuto)

1. Cambia la capacidad máxima en **B12** de 1200 a **620**.
2. Verifica que los estados en la fila 15 cambien a **"ALERTA"** para los meses con mayor volumen (Meses 4, 5 y 6) y a **"NORMAL"** o **"SOBRECARGA"** según corresponda.
3. Verifica que los colores del formato condicional se actualicen automáticamente.
4. Restaura el valor original: B12 = 1200.

**Resultado esperado:** El modelo se comporta como un simulador dinámico de capacidad, permitiendo evaluar escenarios de "¿qué pasaría si...?" simplemente modificando el parámetro de capacidad.

### Prueba 3 — Verificación de integridad de fórmulas (1 minuto)

1. Selecciona la celda **F4** y presiona **F2** para entrar en modo edición.
2. Verifica visualmente que las referencias de parámetros ($C4, $D4, $E4) estén en azul (fijas en columna) y la referencia de demanda en verde (relativa en columna para cambiar por mes).
3. Presiona **Escape** sin modificar.
4. Repite para la celda **G7** (SKU-04, Mes 2) y verifica que los parámetros apunten a la fila 7 ($C7, $D7, $E7) y la demanda al Mes 2.

**Resultado esperado:** Las referencias absolutas y relativas están correctamente configuradas en toda la tabla, garantizando que el modelo escale correctamente para cualquier combinación de producto y mes.

### Checklist final de entregable

Antes de entregar tu archivo, verifica que contenga:

- [ ] Tabla de líneas de picking completa (5 productos × 6 meses) con fórmulas funcionales.
- [ ] Totales por mes en la fila de TOTAL LÍNEAS/MES.
- [ ] Gráfico de barras apiladas con título, leyenda y etiquetas correctas.
- [ ] Sección de análisis de capacidad con porcentajes de utilización y estados de alerta.
- [ ] Formato condicional aplicado correctamente (colores de estado y escala de calor).
- [ ] Mínimo 4 recomendaciones operativas redactadas con respaldo cuantitativo.
- [ ] Nota metodológica al final del documento.
- [ ] Archivo guardado como `Lab18_Picking_Capacidad_[TuNombre].xlsx`.

---

## 8. Resolución de Problemas

### Problema 1 — Las fórmulas de líneas de picking muestran valores incorrectos o el error `#¡DIV/0!`

**Síntoma:** Las celdas del rango F4:K8 muestran `#¡DIV/0!`, valores de cero, o números muy diferentes a los valores de referencia de la Tabla de verificación del Paso 2.

**Causa probable:** El error `#¡DIV/0!` ocurre cuando la celda referenciada como "Unidades por Línea" (columna C) está vacía o contiene cero. También puede ocurrir si la referencia a la hoja `Pronóstico_Demanda` está mal escrita (nombre de hoja incorrecto, espacio extra, o la hoja fue renombrada) y Excel no puede recuperar el valor de demanda, devolviendo 0, lo que genera división por cero o resultados inesperados.

**Solución paso a paso:**

1. Haz clic en la celda problemática (por ejemplo, F4) y presiona **F2** para ver la fórmula.
2. Verifica que la referencia a la hoja de pronóstico tenga el formato correcto: `Pronóstico_Demanda!B2`. Si el nombre de la hoja tiene espacios, debe ir entre apóstrofes: `'Pronóstico Demanda'!B2`.
3. Verifica que las celdas C4:C8 contengan los valores numéricos correctos (12, 20, 24, 15, 18). Si están vacías, ingresa los valores manualmente.
4. Si prefieres evitar referencias entre hojas, copia los valores de pronóstico directamente en la hoja `Picking_Model` en un rango auxiliar (por ejemplo, columnas M a S, filas 4 a 8) y ajusta la fórmula para referenciar ese rango local.
5. Después de corregir, usa **Ctrl + Z** para deshacer si realizaste cambios accidentales, o re-escribe la fórmula desde cero en F4 y vuelve a arrastrarla.

---

### Problema 2 — El gráfico de barras apiladas no muestra los 5 productos como series separadas, o los meses aparecen como series en lugar de como eje horizontal

**Síntoma:** El gráfico muestra los meses (Mes 1 a Mes 6) como series de colores en la leyenda y los productos como etiquetas del eje horizontal, cuando debería ser al revés. O bien, aparecen solo 1 o 2 series en lugar de 5.

**Causa probable:** Excel interpretó la orientación de los datos de forma inversa (filas como series en lugar de columnas, o viceversa). Esto ocurre frecuentemente cuando el rango seleccionado tiene más columnas que filas, y Excel asume que las columnas son las categorías. También puede ocurrir si la selección con **Ctrl** no capturó correctamente ambos rangos (B3:B8 y F3:K8).

**Solución paso a paso:**

1. Haz clic sobre el gráfico para seleccionarlo.
2. Ve a la pestaña **Diseño de gráfico** (que aparece en la cinta cuando el gráfico está seleccionado) → haz clic en **Cambiar fila/columna**. Esto invierte la orientación de los datos y generalmente resuelve el problema.
3. Si el problema persiste, haz clic derecho sobre el gráfico → **Seleccionar datos**.
4. En el cuadro de diálogo, verifica:
   - **Entradas de leyenda (Series):** Deben aparecer los 5 SKUs/productos. Si aparecen "Mes 1", "Mes 2", etc., elimina todas las series y agrégalas manualmente haciendo clic en **Agregar** para cada producto, apuntando al rango de datos correspondiente (por ejemplo, para SKU-01: nombre = B4, valores = F4:K4).
   - **Etiquetas del eje horizontal:** Deben mostrar "Mes 1" a "Mes 6". Haz clic en **Editar** en la sección de etiquetas y selecciona el rango F3:K3.
5. Haz clic en **Aceptar** y verifica que el gráfico ahora muestre correctamente los 5 productos apilados en 6 grupos de barras (uno por mes).

---

## 9. Limpieza y Cierre

Al finalizar la práctica, realiza las siguientes acciones de cierre:

1. **Guarda la versión final** del archivo: **Ctrl + S**. Confirma que el nombre sea `Lab18_Picking_Capacidad_[TuNombre].xlsx`.

2. **Restaura todos los valores de prueba** si realizaste cambios durante las pruebas de validación:
   - Celda D6 (% Dev SKU-03): debe ser `0.04`
   - Celda B12 (Capacidad máxima): debe ser `1200`
   - Celda C4 (Und/Línea SKU-01): debe ser `12`

3. **Verifica el estado final** de la celda F15 (Estado Mes 1): debe mostrar `🟢 HOLGADO` (o el texto equivalente si no usaste emojis).

4. **Cierra el archivo** `Lab18_Pronostico_Demanda.xlsx` original sin guardar cambios (si lo abriste por separado).

5. **Sube o entrega** el archivo `Lab18_Picking_Capacidad_[TuNombre].xlsx` al repositorio o canal indicado por el instructor.

6. **No elimines** el archivo de tu computadora: será utilizado como referencia en las Prácticas 19 y 20, donde integrarás este análisis de capacidad de picking con el cálculo de stock de seguridad y nivel de servicio.

---

## 10. Resumen y Recursos Adicionales

### Resumen de lo aprendido

En esta práctica aplicaste el **puente fundamental entre el pronóstico estadístico y la operación logística**: tomaste números de demanda pronosticada (generados con Promedio Móvil Simple, como aprendiste en la Lección 3.1) y los transformaste en una métrica operativa concreta —las **líneas de picking**— que el equipo de almacén puede usar directamente para planificar turnos, asignar recursos y anticipar cuellos de botella.

Los conceptos clave que aplicaste:

| Concepto | Aplicación en la práctica |
|---|---|
| **Promedio Móvil Simple como base de pronóstico** | Los valores de demanda pronosticada (Tabla A) fueron generados con PM-3, ilustrando la suavidad característica de este método |
| **Fórmula de líneas de picking** | `(Demanda / Und_Línea) × (1 + %Dev) × F.Repick` — integra parámetros logísticos reales |
| **Referencias absolutas en Excel** | `$C4`, `$D4`, `$E4` para fijar parámetros y permitir arrastre de fórmulas |
| **Análisis de capacidad** | Comparación dinámica entre volumen estimado y capacidad máxima con umbrales de alerta |
| **Formato condicional** | Sistema de semáforo visual para identificar riesgos operativos de forma inmediata |
| **Recomendaciones operativas** | Cierre del ciclo: del dato estadístico a la decisión de gestión de almacén |

### Conexión con el contexto del curso

Esta práctica es un ejemplo directo de la pregunta que el instructor debe mantener presente en todas las sesiones: **"¿Cómo impacta este análisis en una decisión operativa de logística?"** La respuesta aquí es clara y cuantificable: el modelo construido permite al jefe de almacén saber, con 6 meses de anticipación, cuántas líneas de picking debe procesar cada mes, qué producto genera mayor carga operativa, y cuándo podría necesitar reforzar el equipo de picking. Eso es forecasting logístico aplicado.

### Recursos adicionales

- **Hyndman, R.J. & Athanasopoulos, G.** — *Forecasting: Principles and Practice*, Capítulo 3: Métodos de suavizamiento. Disponible gratuitamente en: [https://otexts.com/fpp3/](https://otexts.com/fpp3/)
- **Microsoft Support** — Referencia completa de la función `SI` anidada en Excel: [https://support.microsoft.com/es-es/office/función-si](https://support.microsoft.com/es-es/office/función-si-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)
- **Microsoft Support** — Cómo aplicar formato condicional en Excel: [https://support.microsoft.com/es-es/office/usar-formato-condicional](https://support.microsoft.com/es-es/office/usar-el-formato-condicional-para-resaltar-información-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)
- **APICS/ASCM** — Glosario de términos de cadena de suministro (incluye definiciones de nivel de servicio, picking y capacidad operativa): [https://www.ascm.org](https://www.ascm.org/learning-development/certifications-credentials/cscp/)
- **Chopra, S. & Meindl, P.** — *Supply Chain Management: Strategy, Planning, and Operation*, Capítulo 11: Gestión de inventarios en la cadena de suministro. (Pearson)

---

> **🔗 Próxima práctica — Práctica 19:** Utilizarás el modelo de líneas de picking construido en esta práctica como insumo para calcular el **stock de seguridad** necesario para mantener el nivel de servicio objetivo, integrando la variabilidad del pronóstico (desviación estándar del error) con los requerimientos operativos del almacén.

---
LAB_END---
