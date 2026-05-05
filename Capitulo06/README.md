# Análisis de datos históricos utilizando IA generativa

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | 6 — Uso de IA Generativa en Análisis de Demanda Logística |
| **Herramientas principales** | Microsoft Copilot (copilot.microsoft.com), Microsoft Excel |
| **Archivo de trabajo** | `Lab06_Dataset_Anonimizado.xlsx` (proporcionado por el instructor) |

---

## 2. Descripción General

En esta práctica aplicarás directamente los conceptos de la Lección 6.1 interactuando con Microsoft Copilot para analizar datos históricos de demanda de un caso logístico real anonimizado. Comenzarás revisando un protocolo de seguridad y privacidad de datos antes de construir una secuencia de cuatro prompts progresivamente más específicos, siguiendo la estructura **Rol + Contexto + Tarea + Formato**. Finalmente, documentarás y evaluarás las respuestas obtenidas en una tabla de seguimiento en Excel, comparando las observaciones de Copilot con el análisis estadístico que realizaste en módulos anteriores.

> **Anclaje logístico:** Cada prompt que construyas debe responder a una pregunta operativa concreta: ¿qué producto tiene mayor riesgo de quiebre de stock? ¿Qué mes presenta mayor variabilidad? ¿Cómo impacta el coeficiente de variación en el nivel de servicio?

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica, serás capaz de:

- [ ] **Estructurar prompts efectivos** para consultar a Copilot sobre patrones de demanda histórica, aplicando la técnica Rol + Contexto + Tarea + Formato para obtener análisis logísticos relevantes.
- [ ] **Utilizar Copilot para interpretar variabilidad de demanda** e identificar riesgos de inventario a partir de datos históricos proporcionados en el prompt, comparando la interpretación de la IA con el análisis propio.
- [ ] **Aplicar criterios de seguridad y privacidad de datos** al interactuar con herramientas de IA generativa, identificando qué información puede compartirse y qué debe anonimizarse antes de incluirla en un prompt.

---

## 4. Prerrequisitos

### Conocimiento previo requerido
- Haber completado las Prácticas 18–21 (análisis estadístico de demanda: promedio, mediana, desviación estándar, coeficiente de variación y clasificación de productos).
- Haber leído el material de la Lección 6.1 sobre qué es la IA generativa, qué es Microsoft Copilot y cuáles son sus capacidades y limitaciones.
- Comprender los conceptos de variabilidad de demanda, coeficiente de variación (CV) y riesgo de quiebre de stock en contexto logístico.

### Acceso requerido
- Cuenta Microsoft activa con acceso a **copilot.microsoft.com** (cuenta personal o corporativa).
- Microsoft Excel 2016 o superior (Microsoft 365 recomendado).
- Archivo `Lab06_Dataset_Anonimizado.xlsx` distribuido por el instructor.
- Navegador web actualizado (Microsoft Edge o Google Chrome 2024).

> **⚠️ Verificación previa:** Antes de iniciar, abre [copilot.microsoft.com](https://copilot.microsoft.com) y confirma que puedes escribir un mensaje y recibir respuesta. Si no tienes acceso, comunícate con el instructor — se dispone de alternativas (ChatGPT, Gemini) con prompts equivalentes.

---

## 5. Entorno de Laboratorio

### Hardware mínimo requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | i7 / Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1280×768 | 1920×1080 |
| Conexión a internet | 10 Mbps | 20 Mbps |

### Software requerido

| Software | Versión | Uso en esta práctica |
|---|---|---|
| Microsoft Excel | 2016 o superior | Tabla de seguimiento de prompts y evaluación |
| Microsoft Copilot | Versión web (copilot.microsoft.com) | Análisis de demanda mediante prompts |
| Microsoft Edge / Chrome | Actualizado 2024 | Acceso a Copilot web |

### Configuración inicial (antes de comenzar el cronómetro)

```
1. Abre Excel y carga el archivo: Lab06_Dataset_Anonimizado.xlsx
2. Abre una nueva pestaña del navegador y navega a: https://copilot.microsoft.com
3. Inicia sesión con tu cuenta Microsoft
4. Verifica que Copilot responda escribiendo: "Hola, estoy listo para trabajar"
5. Coloca Excel y el navegador en ventanas lado a lado (Alt+Tab para alternar)
```

---

## 6. Desarrollo Paso a Paso

---

### Paso 1 — Revisión del protocolo de seguridad y privacidad de datos

⏱️ **Tiempo estimado:** 3 minutos

**Objetivo:** Identificar qué información del dataset NO debe compartirse con herramientas de IA generativa y aplicar el proceso de anonimización antes de construir cualquier prompt.

#### Instrucciones

**1.1** Abre el archivo `Lab06_Dataset_Anonimizado.xlsx` y dirígete a la hoja **"Seguridad_Privacidad"**. Encontrarás una lista de verificación de 5 puntos. Lee cada punto detenidamente:

```
LISTA DE VERIFICACIÓN — DATOS QUE NO DEBEN INCLUIRSE EN PROMPTS DE IA

□ 1. Nombres reales de clientes o cuentas clave
       Ejemplo prohibido: "Cliente: Supermercados La Colonia"
       Versión anonimizada: "Cliente_A" o "Canal_Retail_01"

□ 2. Precios unitarios reales o márgenes de contribución
       Ejemplo prohibido: "Precio de venta: $47.50 / unidad"
       Versión anonimizada: omitir precios o usar índices relativos

□ 3. Datos personales de empleados o proveedores
       Ejemplo prohibido: "Proveedor: Juan García, RUC 20512345678"
       Versión anonimizada: "Proveedor_1" o "Supplier_01"

□ 4. Información financiera sensible (costos, presupuestos, pérdidas)
       Ejemplo prohibido: "Pérdida por quiebre de stock: $12,400 en Q3"
       Versión anonimizada: usar porcentajes o rangos cualitativos

□ 5. Datos de sistemas internos con identificadores únicos
       Ejemplo prohibido: "SAP Material ID: 4500123456"
       Versión anonimizada: "SKU_A1" hasta "SKU_H8"
```

**1.2** Ahora dirígete a la hoja **"Dataset_Original"** del archivo. Observa los nombres de columna originales. Compáralos con la hoja **"Dataset_Anonimizado"** donde el instructor ya aplicó la anonimización básica. Verifica que se cumplan los 5 puntos de la lista.

**1.3** En la hoja **"Dataset_Anonimizado"**, localiza la tabla resumen que usarás en los prompts. Tiene el siguiente formato (valores de referencia del caso de estudio):

| SKU | Promedio Mensual (uds) | Desv. Estándar | CV (%) | Mín | Máx |
|---|---|---|---|---|---|
| Producto_A | 1,240 | 186 | 15.0% | 890 | 1,580 |
| Producto_B | 876 | 315 | 36.0% | 320 | 1,420 |
| Producto_C | 2,105 | 168 | 8.0% | 1,820 | 2,390 |
| Producto_D | 445 | 222 | 49.9% | 120 | 890 |
| Producto_E | 1,890 | 567 | 30.0% | 980 | 2,870 |
| Producto_F | 3,210 | 193 | 6.0% | 2,890 | 3,540 |
| Producto_G | 678 | 339 | 50.0% | 180 | 1,240 |
| Producto_H | 1,456 | 291 | 20.0% | 980 | 1,920 |

> **Nota:** Si tu archivo contiene valores diferentes, utiliza los datos reales de tu archivo. La tabla anterior es referencial para que puedas verificar el orden de magnitud.

**1.4** Copia esta tabla (selecciona el rango, Ctrl+C) para tenerla lista para pegar en los prompts.

#### Resultado esperado
- Lista de verificación revisada y comprendida.
- Dataset anonimizado identificado y tabla resumen copiada al portapapeles.
- Comprensión clara de por qué cada categoría de dato debe protegerse.

#### Verificación
> Pregúntate: ¿Podría alguien identificar a mi empresa o clientes reales con los datos que voy a compartir? Si la respuesta es "sí" o "tal vez", aplica más anonimización antes de continuar.

---

### Paso 2 — Construcción de la tabla de seguimiento en Excel

⏱️ **Tiempo estimado:** 1 minuto

**Objetivo:** Preparar la estructura de documentación en Excel donde registrarás cada prompt, la respuesta de Copilot y tu evaluación personal.

#### Instrucciones

**2.1** En el archivo `Lab06_Dataset_Anonimizado.xlsx`, dirígete a la hoja **"Seguimiento_Prompts"** (si no existe, crea una hoja nueva con ese nombre).

**2.2** Construye o verifica que exista la siguiente tabla a partir de la celda A1:

```
Columna A: N° Prompt       (valores: 1, 2, 3, 4)
Columna B: Objetivo        (descripción breve del propósito del prompt)
Columna C: Prompt Enviado  (texto completo del prompt)
Columna D: Respuesta Copilot (resumen de la respuesta, máx. 3 líneas)
Columna E: Calidad (1-5)   (escala numérica)
Columna F: Comentario      (observación personal sobre utilidad/precisión)
Columna G: ¿Coincide con análisis previo? (Sí / Parcialmente / No)
```

**2.3** Aplica formato de tabla (Ctrl+T) para facilitar la lectura. Ajusta el ancho de las columnas C y D a 45 caracteres (aproximadamente).

#### Resultado esperado
- Tabla de seguimiento lista con 7 columnas y formato aplicado.
- Filas 2 a 5 vacías, listas para recibir los datos de los 4 prompts.

#### Verificación
Confirma que la hoja **"Seguimiento_Prompts"** es visible en las pestañas de Excel y que la tabla tiene encabezados en negrita.

---

### Paso 3 — Prompt 1: Análisis descriptivo de patrones de demanda

⏱️ **Tiempo estimado:** 2 minutos

**Objetivo:** Construir y enviar el primer prompt siguiendo la estructura Rol + Contexto + Tarea + Formato, solicitando un análisis descriptivo de los patrones de demanda del dataset.

#### Instrucciones

**3.1** Abre Copilot en el navegador. Asegúrate de estar en una conversación nueva (botón "Nueva conversación" o equivalente).

**3.2** Construye el **Prompt 1** usando la siguiente estructura. Copia el texto, reemplaza `[PEGAR TABLA AQUÍ]` con la tabla que copiaste en el Paso 1.4, y envíalo a Copilot:

```
Actúa como un especialista en análisis de demanda logística con experiencia 
en gestión de inventarios para empresas de consumo masivo.

Tengo el siguiente resumen estadístico de demanda mensual para 8 productos 
de un centro de distribución, con 24 meses de historial histórico. 
Los datos han sido anonimizados:

[PEGAR TABLA AQUÍ]

Analiza estos datos y describe los patrones generales de demanda que observas. 
Identifica qué productos tienen demanda estable, cuáles muestran alta variabilidad 
y si hay algún patrón notable en el conjunto.

Responde en formato de párrafo narrativo, con un máximo de 150 palabras, 
usando lenguaje apropiado para un informe de planificación logística.
```

**3.3** Lee la respuesta de Copilot. Mientras la lees, anota mentalmente:
- ¿Identificó correctamente los productos con mayor y menor variabilidad?
- ¿Usó terminología logística apropiada?
- ¿La respuesta es coherente con lo que tú calculaste en las Prácticas 18–21?

**3.4** Regresa a Excel y completa la **fila 2** de la tabla de seguimiento:
- **Columna B:** "Análisis descriptivo de patrones de demanda"
- **Columna C:** Pega el prompt enviado (versión sin la tabla de datos para ahorrar espacio)
- **Columna D:** Resume la respuesta de Copilot en máximo 3 líneas
- **Columna E:** Asigna una calificación del 1 al 5 (1=inútil, 5=excelente)
- **Columna F:** Escribe un comentario breve sobre la calidad de la respuesta
- **Columna G:** Indica si coincide con tu análisis previo

#### Resultado esperado
Copilot debe generar un párrafo narrativo que identifique a **Producto_F** y **Producto_C** como los más estables (CV < 10%), y a **Producto_D** y **Producto_G** como los de mayor variabilidad (CV ≈ 50%). Si la respuesta no menciona estas diferencias, es una señal de que el prompt necesita más contexto — lo ajustarás en prompts posteriores.

#### Verificación
Confirma que la fila 2 de tu tabla de seguimiento está completamente llena antes de avanzar al Paso 4.

---

### Paso 4 — Prompt 2: Identificación de meses con mayor variabilidad

⏱️ **Tiempo estimado:** 2 minutos

**Objetivo:** Construir un prompt más específico que solicite a Copilot identificar los periodos de mayor riesgo operativo a partir de los rangos de demanda (mínimo y máximo).

#### Instrucciones

**4.1** En la misma conversación de Copilot (para mantener el contexto), construye y envía el **Prompt 2**:

```
Basándote en los datos que te compartí anteriormente, ahora necesito 
que analices los rangos de demanda (valores mínimo y máximo) de cada producto.

Identifica:
1. Qué productos presentan la mayor diferencia entre su demanda máxima 
   y mínima (en términos absolutos y porcentuales respecto al promedio).
2. Qué implicaciones operativas tiene esta variabilidad para la planificación 
   de inventario y el nivel de servicio al cliente.
3. En qué productos el riesgo de sobrestock y el riesgo de quiebre de stock 
   son simultáneamente altos.

Responde en formato de lista numerada, con un máximo de 200 palabras.
```

**4.2** Lee la respuesta. Presta atención especial a si Copilot:
- Calculó correctamente las diferencias entre máximo y mínimo (o solo las describió cualitativamente).
- Vinculó la variabilidad con consecuencias operativas concretas (nivel de servicio, stock de seguridad).
- Identificó correctamente a **Producto_E** como un caso de alta variabilidad absoluta (rango: 980 a 2,870 = diferencia de 1,890 unidades).

> **Recuerda la limitación clave de la Lección 6.1:** Copilot puede equivocarse en cálculos numéricos exactos. Si indica un valor diferente al que tú calculaste en Excel, confía en tu hoja de cálculo y anótalo en el comentario de seguimiento.

**4.3** Completa la **fila 3** de la tabla de seguimiento en Excel con los datos de este prompt y su respuesta.

#### Resultado esperado
Copilot debe generar una lista numerada que vincule la variabilidad de demanda con riesgos operativos específicos. La respuesta de mayor calidad incluirá menciones a **stock de seguridad**, **nivel de servicio** o **fill rate** como consecuencias de la variabilidad observada.

#### Verificación
Verifica que en tu tabla de seguimiento la columna G (¿Coincide con análisis previo?) refleja honestamente si la respuesta de Copilot fue consistente con los cálculos que realizaste en las Prácticas 18–21.

---

### Paso 5 — Prompt 3: Interpretación del coeficiente de variación y riesgo de inventario

⏱️ **Tiempo estimado:** 2 minutos

**Objetivo:** Solicitar a Copilot una interpretación del coeficiente de variación (CV) en términos de clasificación de riesgo de inventario, conectando el indicador estadístico con decisiones operativas de reabastecimiento.

#### Instrucciones

**5.1** Construye y envía el **Prompt 3** en la misma conversación:

```
Actúa como consultor de supply chain. A partir de los coeficientes de variación 
(CV) de los 8 productos que te compartí, necesito que:

1. Clasifiques cada producto en una de las siguientes categorías de riesgo 
   de inventario:
   - BAJO RIESGO: CV < 20%
   - RIESGO MEDIO: CV entre 20% y 40%
   - ALTO RIESGO: CV > 40%

2. Expliques, en términos logísticos, qué tipo de política de reabastecimiento 
   es más adecuada para cada categoría (por ejemplo: reabastecimiento periódico 
   fijo, punto de reorden dinámico, revisión continua).

3. Indiques cuál de las tres categorías requiere mayor stock de seguridad 
   y por qué.

Responde en formato de tabla para la clasificación y un párrafo corto 
(máximo 80 palabras) para la explicación de políticas.
```

**5.2** Verifica la clasificación que genera Copilot contra la siguiente referencia esperada:

| Categoría | Productos esperados |
|---|---|
| BAJO RIESGO (CV < 20%) | Producto_A (15%), Producto_C (8%), Producto_F (6%), Producto_H (20%) |
| RIESGO MEDIO (CV 20–40%) | Producto_B (36%), Producto_E (30%) |
| ALTO RIESGO (CV > 40%) | Producto_D (49.9%), Producto_G (50%) |

> **Nota:** Producto_H tiene CV exactamente del 20%. Dependiendo de cómo Copilot interprete el límite, puede clasificarlo en BAJO o MEDIO RIESGO. Ambas respuestas son válidas — lo importante es que justifique su criterio.

**5.3** Si Copilot comete un error de clasificación (por ejemplo, coloca un producto en la categoría incorrecta), documenta el error en la columna F de tu tabla de seguimiento. Esto es un ejercicio de pensamiento crítico, no un fallo de la práctica.

**5.4** Completa la **fila 4** de la tabla de seguimiento.

#### Resultado esperado
Una tabla de clasificación de 8 productos en 3 categorías de riesgo, más un párrafo que vincule el nivel de riesgo con políticas de reabastecimiento específicas. La respuesta de mayor calidad mencionará que los productos de ALTO RIESGO requieren mayor stock de seguridad debido a la mayor incertidumbre en la demanda.

#### Verificación
Compara la tabla generada por Copilot con tu propia clasificación de las Prácticas 18–21. Registra coincidencias y discrepancias en la columna G.

---

### Paso 6 — Prompt 4: Identificación de productos con mayor riesgo de quiebre de stock

⏱️ **Tiempo estimado:** 1 minuto

**Objetivo:** Construir el prompt más específico y orientado a decisión operativa de la secuencia, solicitando a Copilot una recomendación priorizada de los productos que requieren atención inmediata en la planificación de inventario.

#### Instrucciones

**6.1** Construye y envía el **Prompt 4** — el más orientado a acción operativa de la secuencia:

```
Con base en todo el análisis anterior, necesito una recomendación operativa 
para el equipo de planificación de inventario.

Identifica los 3 productos con mayor riesgo de quiebre de stock y para cada uno:
1. Justifica por qué tiene alto riesgo (usando los indicadores estadísticos).
2. Recomienda una acción operativa concreta y específica (no genérica).
3. Indica qué indicador deberías monitorear semanalmente para este producto.

Responde en formato de ficha por producto, con encabezado, justificación, 
acción recomendada e indicador de monitoreo. Usa lenguaje directo y ejecutivo.
```

**6.2** Evalúa la respuesta de Copilot considerando:
- ¿Priorizó correctamente a los productos de mayor CV (Producto_D, Producto_G, Producto_E o Producto_B)?
- ¿Las acciones recomendadas son específicas y operacionalmente viables, o son genéricas?
- ¿Los indicadores de monitoreo propuestos son relevantes (fill rate, días de cobertura, rotación)?

**6.3** Completa la **fila 5** de la tabla de seguimiento con los datos del Prompt 4.

#### Resultado esperado
Tres fichas de producto con estructura clara: justificación estadística + acción operativa + indicador de monitoreo. Los productos prioritarios esperados son **Producto_D** y **Producto_G** (CV ~50%) y **Producto_B** o **Producto_E** (CV 30–36%) como tercer lugar.

#### Verificación
Pregúntate: ¿Las recomendaciones de Copilot son accionables para un equipo de planificación real? ¿Agregarías o modificarías algo? Anota tu respuesta en la columna F.

---

## 7. Validación y Cierre del Análisis

⏱️ **Tiempo estimado:** Incluido en los pasos anteriores (revisión final: 30 segundos)**

### Lista de verificación de completitud

Antes de cerrar la práctica, confirma que tu tabla de seguimiento en Excel cumple con todos los criterios:

| Criterio | ¿Cumplido? |
|---|---|
| Las 4 filas de la tabla de seguimiento están completamente llenas | ☐ |
| Todos los prompts siguieron la estructura Rol + Contexto + Tarea + Formato | ☐ |
| La columna E (calidad 1–5) tiene valores numéricos en las 4 filas | ☐ |
| La columna G refleja una comparación honesta con el análisis propio | ☐ |
| El dataset utilizado en los prompts corresponde a la versión anonimizada | ☐ |
| Se documentó al menos un error o limitación de Copilot observado | ☐ |

### Pregunta de reflexión final

Responde brevemente (2–3 líneas) en la celda **I1** de la hoja "Seguimiento_Prompts":

> *"¿En qué aspecto del análisis de demanda Copilot fue más útil hoy, y en qué aspecto demostró una limitación que tú como analista debes compensar con tu propio criterio?"*

Esta reflexión conecta directamente con la advertencia de la Lección 6.1: *"Copilot puede equivocarse en cálculos numéricos exactos. Usa Excel para los cálculos; Copilot para la interpretación."*

---

## 8. Solución de Problemas

### Problema 1: Copilot no recuerda los datos de la tabla en el Prompt 2, 3 o 4

**Síntoma:** A partir del Prompt 2, Copilot responde con frases como "No tengo acceso a los datos anteriores" o genera una respuesta genérica sin hacer referencia a los productos específicos del caso de estudio.

**Causa:** Copilot tiene un límite de contexto de conversación. Si la tabla de datos es extensa o si ha pasado tiempo entre prompts, el modelo puede perder el contexto previo. Esto ilustra directamente la limitación mencionada en la Lección 6.1: *"No tiene memoria entre sesiones y puede perder contexto en conversaciones largas."*

**Solución:**
```
Opción A (recomendada): Al inicio de cada prompt posterior, agrega la frase:
"Recuerda que estamos trabajando con el siguiente dataset anonimizado: [pega 
la tabla nuevamente]" antes de formular la tarea específica.

Opción B: Inicia una nueva conversación y en el Prompt 1 incluye TODOS 
los datos y el contexto completo del caso, luego continúa desde ese punto.

Opción C: Usa el resumen de la respuesta anterior como contexto:
"Basándote en tu análisis anterior donde identificaste que Producto_D y 
Producto_G tienen CV ~50%..."
```

---

### Problema 2: Copilot genera clasificaciones de CV incorrectas o cálculos numéricos erróneos

**Síntoma:** En el Prompt 3, Copilot clasifica un producto en una categoría de riesgo incorrecta (por ejemplo, coloca Producto_F con CV=6% en RIESGO MEDIO), o en el Prompt 2 calcula diferencias entre máximo y mínimo con valores que no coinciden con los datos del dataset.

**Causa:** Este es un ejemplo directo del fenómeno de "alucinación" descrito en la Lección 6.1. Los modelos de lenguaje no son calculadoras — razonan sobre números pero pueden cometer errores aritméticos o de clasificación, especialmente cuando hay múltiples valores numéricos en el contexto.

**Solución:**
```
Paso 1: NO corrijas el error dentro de Copilot esperando que "aprenda". 
        Documenta el error en la columna F de tu tabla de seguimiento.

Paso 2: Verifica el cálculo correcto en Excel usando tus datos de las 
        Prácticas 18-21. Tus cálculos en Excel son la fuente de verdad.

Paso 3: Si necesitas que Copilot corrija su respuesta, reformula el prompt 
        con los valores explícitos:
        "El CV de Producto_F es 6.0% (calculado como 193/3210). 
        Con el criterio CV < 20% = BAJO RIESGO, ¿en qué categoría 
        corresponde clasificarlo?"

Paso 4: Registra este hallazgo en tu reflexión final (celda I1) como 
        evidencia de la limitación de la IA en cálculos numéricos exactos.
```

---

## 9. Limpieza y Cierre

### Acciones de cierre requeridas

**9.1** Guarda el archivo Excel con la tabla de seguimiento completa:
```
Nombre sugerido: Lab06_Seguimiento_Prompts_[TuNombre].xlsx
Ubicación: Carpeta del curso / Módulo 6
Formato: .xlsx (Excel Workbook)
```

**9.2** En Copilot, cierra la conversación activa (no es necesario eliminarla, pero evita dejar datos del caso de estudio en conversaciones abiertas en dispositivos compartidos).

**9.3** Si utilizaste una cuenta personal de Microsoft para acceder a Copilot en un equipo compartido o de laboratorio, cierra sesión:
```
1. Haz clic en tu foto de perfil (esquina superior derecha de copilot.microsoft.com)
2. Selecciona "Cerrar sesión"
3. Cierra la pestaña del navegador
```

**9.4** Verifica que no hayas guardado datos sensibles del caso de estudio en el historial del navegador. En equipos compartidos:
```
Chrome/Edge: Ctrl+Shift+Delete → Selecciona "Historial de navegación" 
y "Cookies y datos de sitios" → Eliminar
```

> **Buena práctica de seguridad:** Este paso de cierre de sesión y limpieza de historial es parte del protocolo de uso responsable de IA generativa en entornos empresariales — el mismo principio que revisaste en la lista de verificación del Paso 1.

---

## 10. Resumen y Recursos Adicionales

### Lo que practicaste hoy

En esta práctica aplicaste los conceptos centrales de la Lección 6.1 en un contexto logístico real:

| Concepto aplicado | Cómo lo practicaste |
|---|---|
| **Protocolo de seguridad y privacidad** | Lista de verificación de 5 puntos + anonimización del dataset antes de cualquier prompt |
| **Estructura Rol + Contexto + Tarea + Formato** | 4 prompts progresivamente más específicos con esta estructura explícita |
| **Capacidades de Copilot: interpretación de datos** | Análisis descriptivo, clasificación por CV, identificación de riesgos |
| **Limitaciones de Copilot: alucinaciones numéricas** | Verificación de clasificaciones contra cálculos propios de Excel |
| **Pensamiento crítico sobre respuestas de IA** | Tabla de seguimiento con evaluación 1–5 y comparación con análisis previo |

### Conexión con el ciclo de forecasting logístico

La estructura de los 4 prompts que construiste sigue una progresión lógica que refleja el ciclo de análisis de demanda:

```
Prompt 1 → Describir (¿qué veo en los datos?)
Prompt 2 → Analizar (¿dónde está la variabilidad?)
Prompt 3 → Clasificar (¿qué nivel de riesgo representa?)
Prompt 4 → Decidir (¿qué acciones operativas tomar?)
```

Esta progresión —de descriptivo a prescriptivo— es exactamente la que usarás en las Prácticas 27 y 28, donde generarás escenarios predictivos y recomendaciones automatizadas de reabastecimiento con Copilot.

### Pregunta para el próximo módulo

> *"Si Copilot puede interpretar datos históricos y clasificar riesgos, ¿puede también generar un pronóstico de demanda para los próximos 3 meses? ¿Qué información adicional necesitaría para hacerlo con mayor precisión?"*

Responderás esta pregunta en la **Práctica 27: Generación de escenarios predictivos con IA generativa**.

### Recursos de referencia

- [Microsoft Copilot — Página oficial](https://www.microsoft.com/es-es/microsoft-copilot)
- [Microsoft Learn: Introducción a Copilot para Microsoft 365](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview)
- [Microsoft Learn: Uso de Copilot en Excel para análisis de datos](https://learn.microsoft.com/es-es/copilot/microsoft-365/use-copilot-in-excel)
- [Guía de prompts efectivos para análisis de datos — Microsoft](https://adoption.microsoft.com/es-es/copilot/)
- Material del curso: Lección 6.2 — Estructuración de prompts para análisis logístico (próxima sesión)

---

> **Nota para el instructor:** Si los participantes terminan antes del tiempo asignado, proponga este desafío de extensión: *"Construye un quinto prompt que solicite a Copilot comparar los productos de ALTO RIESGO con los de BAJO RIESGO y genere un párrafo de recomendación ejecutiva para la gerencia de supply chain."* Este ejercicio refuerza la habilidad de escalar el nivel de abstracción en los prompts.

---

---

# Práctica 27 — Comparación de resultados entre Excel y herramientas de IA

## Metadatos

| Campo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Aplicar (Apply) |
| **Práctica número** | 27 |
| **Módulo** | 6 — IA Generativa en Análisis de Demanda |
| **Archivo de entrada** | `Practica21_Pronostico_Seleccionado.xlsx` (generado en Práctica 21) |
| **Archivo de salida** | `Practica27_Comparacion_Excel_IA.xlsx` |

---

## Descripción General

En esta práctica aplicarás pensamiento crítico comparativo para evaluar las diferencias entre el pronóstico estadístico calculado en Excel (Práctica 21) y los escenarios predictivos generados por Copilot mediante prompts estructurados. Construirás en Excel una tabla de comparación multi-escenario, un gráfico de líneas con cuatro series simultáneas y un mini-framework de decisión que te permita determinar operativamente cuándo confiar en cada herramienta. El resultado es un criterio profesional documentado sobre el uso complementario de estadística y IA generativa en contextos de planificación logística.

---

## Objetivos de Aprendizaje

- [ ] Generar tres escenarios predictivos (optimista, conservador y crítico) mediante un prompt estructurado en Copilot, obteniendo valores numéricos específicos para el próximo trimestre.
- [ ] Comparar sistemáticamente los escenarios de Copilot contra el pronóstico de Excel, calculando diferencias porcentuales e identificando convergencias y divergencias.
- [ ] Desarrollar y documentar un criterio crítico sobre las condiciones de uso preferente de cada herramienta, plasmado en un mini-framework de decisión en Excel.

---

## Prerrequisitos

### Conocimiento previo
- Haber completado la **Práctica 21** y contar con el pronóstico seleccionado (modelo con menor MAPE) con valores numéricos disponibles para los meses M19, M20 y M21.
- Haber completado la **Práctica 26** y tener experiencia básica en la construcción y ejecución de prompts en Copilot.
- Comprensión de los conceptos de escenario optimista (+10–15%), conservador (±5%) y crítico (–10–20%) en planificación logística de inventario.
- Familiaridad con los conceptos de IA generativa presentados en la **Lección 6.1**: capacidades, limitaciones y riesgo de alucinación.

### Acceso y recursos
- Acceso activo a **Microsoft Copilot** (copilot.microsoft.com) o herramienta equivalente (ChatGPT, Gemini).
- Archivo `Practica21_Pronostico_Seleccionado.xlsx` abierto y disponible.
- **Microsoft Excel 2016 o superior** (Microsoft 365 recomendado).
- Conexión a internet estable (mínimo 10 Mbps).

---

## Entorno de Laboratorio

### Hardware mínimo requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Core i7 / Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 500 MB | 2 GB |
| Resolución de pantalla | 1280×768 | 1920×1080 |
| Conexión a internet | 10 Mbps | 25 Mbps |

### Software requerido

| Software | Versión | Uso en esta práctica |
|---|---|---|
| Microsoft Excel | 2016 o superior (M365 recomendado) | Tabla comparativa, gráfico, framework |
| Microsoft Copilot / ChatGPT / Gemini | Versión web actualizada | Generación de escenarios predictivos |
| Microsoft Edge / Google Chrome | Versión 2024 | Acceso a herramienta de IA |

### Configuración inicial (antes de comenzar)

1. Abre `Practica21_Pronostico_Seleccionado.xlsx` en Excel. Localiza la hoja con los valores pronosticados para M19, M20 y M21 del SKU de trabajo. Anota los tres valores en papel o en un bloc de notas — los necesitarás en el Paso 2.
2. Abre una pestaña nueva en tu navegador y navega a [https://copilot.microsoft.com](https://copilot.microsoft.com). Inicia sesión si es necesario.
3. Crea un nuevo archivo Excel en blanco y guárdalo como `Practica27_Comparacion_Excel_IA.xlsx`.

> ⚠️ **Nota de tiempo:** Esta práctica tiene 11 minutos. Cada paso indica el tiempo sugerido. Respeta los tiempos para completar todos los entregables.

---

## Pasos del Laboratorio

---

### Paso 1 — Preparar el resumen de historial para el prompt de Copilot

**Objetivo:** Construir el contexto de datos que se incluirá en el prompt, resumiendo 18 meses de historial de demanda del SKU seleccionado en un formato anonimizado apto para compartir con la IA.

⏱️ *Tiempo sugerido: 2 minutos*

**Instrucciones:**

1. En tu archivo `Practica27_Comparacion_Excel_IA.xlsx`, crea una hoja llamada **`Datos_Prompt`**.

2. En la celda **A1** escribe el encabezado: `Mes` y en **B1**: `Demanda_Unidades`. Completa la tabla con los datos de los últimos 18 meses del SKU de trabajo, usando la nomenclatura M1 a M18 (datos anonimizados). La tabla debe verse así:

   ```
   Mes | Demanda_Unidades
   M1  | [valor real del historial]
   M2  | [valor real del historial]
   ...
   M18 | [valor real del historial]
   ```

3. En la celda **D1** escribe: `Pronóstico Excel (modelo seleccionado Práctica 21)` y en **E1**: `Período`. Completa:

   ```
   E2: M19   D2: [valor pronóstico M19 de Práctica 21]
   E3: M20   D3: [valor pronóstico M20 de Práctica 21]
   E4: M21   D4: [valor pronóstico M21 de Práctica 21]
   ```

4. Guarda el archivo (`Ctrl + S`).

**Resultado esperado:** Una hoja con 18 filas de historial anonimizado y 3 filas con los pronósticos de referencia de Excel, listos para ser referenciados en el prompt.

**Verificación:** Confirma que los valores en D2:D4 coinciden exactamente con los pronósticos de tu Práctica 21. Si hay discrepancia, regresa al archivo fuente antes de continuar.

---

### Paso 2 — Construir y ejecutar el prompt estructurado en Copilot

**Objetivo:** Redactar un prompt con estructura completa (contexto + historial + solicitud de tres escenarios + formato de respuesta esperado) y ejecutarlo en Copilot para obtener valores numéricos específicos.

⏱️ *Tiempo sugerido: 3 minutos*

**Instrucciones:**

1. En tu navegador, abre Copilot (o la herramienta de IA disponible). Inicia una **nueva conversación**.

2. Copia el siguiente prompt en el campo de texto, **sustituyendo los valores entre corchetes** con los datos reales de tu hoja `Datos_Prompt`:

   ```
   Actúa como un analista de planificación de demanda logística con experiencia 
   en forecasting de inventario.

   Tengo el siguiente historial de demanda mensual (18 meses) para un producto 
   de consumo masivo en una operación de distribución:

   M1:[val], M2:[val], M3:[val], M4:[val], M5:[val], M6:[val],
   M7:[val], M8:[val], M9:[val], M10:[val], M11:[val], M12:[val],
   M13:[val], M14:[val], M15:[val], M16:[val], M17:[val], M18:[val]

   Necesito que generes tres escenarios de demanda para los próximos 3 meses 
   (M19, M20 y M21) con las siguientes definiciones:

   - ESCENARIO OPTIMISTA: crecimiento del 10% al 15% respecto al promedio 
     reciente (últimos 3 meses)
   - ESCENARIO CONSERVADOR: demanda estable con variación de ±5% respecto 
     al promedio reciente
   - ESCENARIO CRÍTICO: caída del 10% al 20% respecto al promedio reciente

   Por favor, responde ÚNICAMENTE con una tabla en este formato exacto:

   | Período | Escenario_Optimista | Escenario_Conservador | Escenario_Crítico |
   |---------|--------------------|-----------------------|-------------------|
   | M19     | [valor numérico]   | [valor numérico]      | [valor numérico]  |
   | M20     | [valor numérico]   | [valor numérico]      | [valor numérico]  |
   | M21     | [valor numérico]   | [valor numérico]      | [valor numérico]  |

   Después de la tabla, incluye en 2-3 líneas los supuestos clave que usaste 
   para calcular cada escenario.
   ```

3. Envía el prompt y espera la respuesta completa de Copilot.

4. **Copia los 9 valores numéricos** de la tabla de respuesta (3 períodos × 3 escenarios) y los supuestos indicados. Pégalos temporalmente en el Bloc de notas o en la hoja `Datos_Prompt` celdas G1:J5 para no perderlos.

> 💡 **Si Copilot no genera la tabla en el formato solicitado:** Envía un prompt de seguimiento: *"Por favor, presenta los resultados únicamente en la tabla con el formato que te indiqué, con valores numéricos enteros."*

> ⚠️ **Recordatorio Lección 6.1:** Copilot puede "alucinar" valores. Verifica que los números generados sean coherentes con el rango histórico de tu demanda (no deben ser órdenes de magnitud diferentes al historial).

**Resultado esperado:** Una tabla de 3×3 con valores numéricos específicos para los tres escenarios en M19, M20 y M21, más una breve descripción de supuestos.

**Verificación:** Los valores del escenario conservador deben estar aproximadamente en el rango del promedio de los últimos 3 meses del historial ±5%. Si los valores son absurdos (ej.: negativos, o 10× mayores al historial), el prompt debe reejecutarse con datos corregidos.

---

### Paso 3 — Construir la tabla comparativa en Excel

**Objetivo:** Crear la tabla de 6 columnas con todos los valores (Excel + 3 escenarios IA) y calcular la diferencia porcentual entre el escenario conservador de IA y el pronóstico de Excel.

⏱️ *Tiempo sugerido: 2 minutos*

**Instrucciones:**

1. En `Practica27_Comparacion_Excel_IA.xlsx`, crea una nueva hoja llamada **`Comparacion`**.

2. Construye el encabezado en la fila 1 con el siguiente formato exacto:

   | Celda | Contenido |
   |---|---|
   | A1 | `Período` |
   | B1 | `Pronóstico_Excel` |
   | C1 | `Escenario_Optimista_IA` |
   | D1 | `Escenario_Conservador_IA` |
   | E1 | `Escenario_Crítico_IA` |
   | F1 | `Dif%_Conservador_vs_Excel` |

3. Ingresa los datos en las filas 2 a 4:

   ```
   A2: M19   B2: [Pronóstico Excel M19]   C2: [Optimista M19]   D2: [Conservador M19]   E2: [Crítico M19]
   A3: M20   B3: [Pronóstico Excel M20]   C3: [Optimista M20]   D3: [Conservador M20]   E3: [Crítico M20]
   A4: M21   B4: [Pronóstico Excel M21]   C4: [Optimista M21]   D4: [Conservador M21]   E4: [Crítico M21]
   ```

4. En la celda **F2**, ingresa la fórmula de diferencia porcentual:

   ```excel
   =(D2-B2)/B2
   ```

5. Aplica formato de porcentaje a F2 (`Ctrl+Shift+%`) y copia la fórmula hacia abajo a **F3** y **F4**.

6. Aplica formato condicional a F2:F4:
   - Selecciona F2:F4 → `Inicio` → `Formato condicional` → `Escalas de color`
   - Elige la escala Verde-Amarillo-Rojo (verde = convergencia cercana a 0%, rojo = divergencia alta).

7. En la celda **B6** escribe: `Promedio diferencia absoluta:` y en **C6**:

   ```excel
   =PROMEDIO(ABS(F2),ABS(F3),ABS(F4))
   ```

   Aplica formato de porcentaje a C6.

8. Guarda el archivo.

**Resultado esperado:** Una tabla con 4 columnas de valores numéricos, una columna de diferencias porcentuales con formato condicional y un promedio de divergencia general.

**Verificación:** Confirma que las fórmulas en F2:F4 calculan correctamente: si el conservador IA = Excel exacto, el resultado debe ser 0%. Si el conservador IA es 5% mayor, el resultado debe ser 5,0%.

---

### Paso 4 — Crear el gráfico de líneas con cuatro series simultáneas

**Objetivo:** Visualizar en un gráfico de líneas las cuatro series (Pronóstico Excel + 3 escenarios IA) para identificar visualmente convergencias y divergencias entre métodos.

⏱️ *Tiempo sugerido: 2 minutos*

**Instrucciones:**

1. En la hoja **`Comparacion`**, selecciona el rango **A1:E4** (encabezados + datos de las cuatro series, sin la columna de diferencia).

2. Ve a `Insertar` → `Gráficos` → **Gráfico de líneas** → selecciona **"Línea con marcadores"**.

3. Excel insertará el gráfico con cuatro series. Ajusta el formato:

   - **Haz clic derecho sobre el gráfico** → `Seleccionar datos` → verifica que las 4 series tengan nombres correctos: `Pronóstico_Excel`, `Escenario_Optimista_IA`, `Escenario_Conservador_IA`, `Escenario_Crítico_IA`.
   - Si los nombres no aparecen correctamente, edita cada serie en `Seleccionar datos` → `Editar` y asigna el nombre desde la celda de encabezado correspondiente.

4. Aplica el siguiente formato visual para facilitar la lectura:

   | Serie | Color sugerido | Tipo de línea |
   |---|---|---|
   | `Pronóstico_Excel` | Azul oscuro (#003087) | Línea sólida, grosor 2.5 pt |
   | `Escenario_Optimista_IA` | Verde (#00B050) | Línea discontinua |
   | `Escenario_Conservador_IA` | Naranja (#FF6600) | Línea sólida, grosor 1.5 pt |
   | `Escenario_Crítico_IA` | Rojo (#FF0000) | Línea discontinua |

   Para cambiar el color: clic derecho sobre cada línea → `Dar formato a serie de datos` → `Color de línea`.

5. Agrega un título al gráfico: **"Comparación Pronóstico Excel vs. Escenarios IA — Trimestre M19-M21"**

6. Activa las etiquetas de datos: clic derecho sobre cada línea → `Agregar etiquetas de datos`.

7. Mueve el gráfico debajo de la tabla (fila 8 en adelante) y redimensiónalo para que ocupe aproximadamente el rango A8:F22.

8. Guarda el archivo.

**Resultado esperado:** Un gráfico de líneas con cuatro series claramente diferenciadas por color y estilo, con etiquetas de datos visibles, que muestre visualmente si el escenario conservador de IA converge o diverge con el pronóstico de Excel.

**Verificación:** El gráfico debe mostrar exactamente 4 líneas. Si aparecen menos, verifica que el rango de datos incluya todas las columnas B:E. Si el eje X muestra números en lugar de M19/M20/M21, ajusta en `Seleccionar datos` → `Etiquetas del eje horizontal`.

---

### Paso 5 — Responder las 5 preguntas de análisis crítico

**Objetivo:** Desarrollar criterio analítico documentando las reflexiones sobre convergencia, limitaciones y complementariedad de ambas herramientas, directamente en la hoja Excel.

⏱️ *Tiempo sugerido: 1.5 minutos*

**Instrucciones:**

1. En la hoja **`Comparacion`**, a partir de la fila **25**, construye la siguiente sección de análisis crítico. Escribe cada pregunta en la columna A y tu respuesta en la columna B (con ajuste de texto activado):

   | Fila | Columna A — Pregunta | Columna B — Tu respuesta |
   |---|---|---|
   | 25 | **ANÁLISIS CRÍTICO** | *(encabezado, celda combinada A25:F25, fondo azul oscuro, texto blanco)* |
   | 26 | **P1.** ¿El escenario conservador de IA coincide con el pronóstico Excel? ¿Cuál es la diferencia porcentual promedio? | [Escribe tu respuesta basada en los datos de F2:F4 y C6] |
   | 27 | **P2.** ¿Qué factores podría estar considerando Copilot que el modelo estadístico de Excel no captura? | [Ej.: conocimiento de mercado externo, eventos macroeconómicos, comportamiento de categoría] |
   | 28 | **P3.** ¿Cuál de los tres escenarios de IA es más útil para calcular el inventario de seguridad? ¿Por qué? | [Ej.: el escenario crítico define el peor caso; el optimista define la capacidad máxima necesaria] |
   | 29 | **P4.** ¿En qué contextos confiarías más en el pronóstico estadístico de Excel que en Copilot? | [Ej.: datos históricos limpios y estables, sin eventos externos, necesidad de reproducibilidad] |
   | 30 | **P5.** ¿Cómo usarías ambas herramientas de forma complementaria en un ciclo mensual de planificación? | [Ej.: Excel para el pronóstico base; Copilot para escenarios de riesgo y comunicación ejecutiva] |

2. Aplica formato a la sección:
   - Fila 25: combina A25:F25, fondo `#003087`, texto blanco, negrita, tamaño 11.
   - Filas 26–30: columna A con fondo `#D9E1F2` (azul claro), columna B con ajuste de texto activado y altura de fila suficiente para leer las respuestas.

3. Guarda el archivo.

> 📝 **Nota para el participante:** No existe una respuesta "correcta" única. Lo que se evalúa es la coherencia del razonamiento con los datos obtenidos en esta práctica y con los conceptos de la Lección 6.1 (capacidades y limitaciones de la IA generativa).

**Resultado esperado:** Cinco respuestas documentadas directamente en Excel, con formato visual diferenciado, que evidencien pensamiento crítico sobre las diferencias entre ambas herramientas.

**Verificación:** Cada celda de respuesta debe contener al menos 1 oración completa. Las respuestas a P1 deben referenciar explícitamente los valores numéricos calculados en la columna F.

---

### Paso 6 — Construir el mini-framework de decisión (diagrama de flujo)

**Objetivo:** Sintetizar el criterio de uso de cada herramienta en un diagrama de flujo visual construido con formas de Excel (SmartArt o formas básicas), que pueda ser reutilizado en el trabajo diario.

⏱️ *Tiempo sugerido: 0.5 minutos (construcción rápida)*

**Instrucciones:**

1. Crea una nueva hoja en el archivo y nómbrala **`Framework_Decision`**.

2. Ve a `Insertar` → `Ilustraciones` → `SmartArt` → selecciona la categoría **"Proceso"** → elige **"Proceso básico"** (flechas en línea horizontal).

   > **Alternativa si SmartArt no está disponible:** Usa `Insertar` → `Ilustraciones` → `Formas` y construye el diagrama manualmente con rectángulos y flechas.

3. Construye el siguiente flujo de decisión con 5 nodos. Edita el texto de cada forma del SmartArt haciendo doble clic:

   ```
   [NODO 1]
   ¿Tengo datos
   históricos limpios
   (>12 meses)?
   
        SÍ ──────────────────────────────────────────────────────────────────────────────────▶
        NO ──────────────────────────────────────────────────────────────────────────────────▶
   ```

   Dado que SmartArt en modo "proceso básico" no admite bifurcaciones, construye el framework como dos rutas paralelas usando formas básicas:

   **Ruta A — Usar Excel (estadístico):**
   - Forma 1 (rombo, amarillo): `¿Datos históricos limpios y estables (>12 meses)?`
   - Flecha → `SÍ`
   - Forma 2 (rectángulo, azul): `USAR EXCEL\nPronóstico estadístico (promedio móvil / suavización / regresión)`
   - Forma 3 (rectángulo, verde claro): `Resultado: valor puntual con error medible (MAE/MAPE)`

   **Ruta B — Usar Copilot (IA generativa):**
   - Forma 4 (rombo, amarillo): `¿Necesito escenarios de riesgo o comunicación ejecutiva?`
   - Flecha → `SÍ`
   - Forma 5 (rectángulo, naranja): `USAR COPILOT\nGeneración de escenarios (optimista / conservador / crítico)`
   - Forma 6 (rectángulo, verde claro): `Resultado: rango de demanda posible para decisiones de inventario de seguridad`

   **Nodo de cierre (ambas rutas convergen):**
   - Forma 7 (rectángulo redondeado, azul oscuro, texto blanco): `USO COMPLEMENTARIO\nExcel define el pronóstico base → Copilot enriquece con escenarios → Decisión operativa integrada`

4. Agrega un título sobre el diagrama: **"¿Cuándo usar Excel vs. Copilot en planificación de demanda?"** (cuadro de texto, fuente 14, negrita).

5. Agrega una nota al pie del diagrama (cuadro de texto pequeño, fuente 9, cursiva):
   ```
   Nota: Copilot no reemplaza el cálculo estadístico. Siempre verifica los 
   valores numéricos generados por IA contra el historial de demanda real. 
   (Ref.: Lección 6.1 — Limitaciones de la IA generativa)
   ```

6. Guarda el archivo (`Ctrl + S`).

**Resultado esperado:** Una hoja con un diagrama de flujo visual que muestra claramente las condiciones de uso preferente de Excel y Copilot, con un nodo de convergencia que enfatiza el uso complementario.

**Verificación:** El diagrama debe contener al menos 2 rombos de decisión (condiciones), 2 rectángulos de acción (herramienta a usar) y 1 nodo de cierre. Todos los nodos deben tener texto legible sin solapamiento.

---

## Validación y Prueba Final

Una vez completados los 6 pasos, realiza la siguiente verificación integral del archivo `Practica27_Comparacion_Excel_IA.xlsx`:

### Lista de verificación de entregables

| # | Entregable | Ubicación | Criterio de aceptación |
|---|---|---|---|
| 1 | Historial de 18 meses anonimizado | Hoja `Datos_Prompt`, A1:B19 | 18 filas de datos + encabezado |
| 2 | Valores pronóstico Excel M19-M21 | Hoja `Datos_Prompt`, D2:E4 | Coinciden con Práctica 21 |
| 3 | Valores escenarios IA (9 celdas) | Hoja `Comparacion`, C2:E4 | Coherentes con rango histórico |
| 4 | Diferencias porcentuales calculadas | Hoja `Comparacion`, F2:F4 | Fórmula `=(D-B)/B` aplicada |
| 5 | Promedio diferencia absoluta | Hoja `Comparacion`, C6 | Fórmula con ABS aplicada |
| 6 | Gráfico de líneas con 4 series | Hoja `Comparacion`, fila 8+ | 4 líneas diferenciadas por color |
| 7 | 5 respuestas de análisis crítico | Hoja `Comparacion`, filas 26-30 | Mínimo 1 oración por respuesta |
| 8 | Mini-framework de decisión | Hoja `Framework_Decision` | Diagrama con 2+ rombos y nodo de cierre |

### Prueba de coherencia numérica

Ejecuta esta verificación rápida: el valor del **Escenario_Conservador_IA** para M19 debe estar dentro del rango:

```
Límite inferior: PROMEDIO(M16:M18) × 0.95
Límite superior: PROMEDIO(M16:M18) × 1.05
```

Si el valor de Copilot está fuera de este rango, documenta la discrepancia en la respuesta a P1 del análisis crítico y explica la posible causa (ej.: Copilot pudo haber detectado tendencia creciente en el historial).

---

## Resolución de Problemas

### Problema 1 — Copilot no genera la tabla en el formato solicitado

**Síntoma:** Copilot responde con texto narrativo, valores sin estructura tabular, o genera una tabla con formato diferente al solicitado (ej.: usa guiones en lugar de barras verticales, o incluye rangos en lugar de valores únicos).

**Causa:** Los modelos de lenguaje tienden a adaptar el formato de respuesta cuando el prompt es ambiguo o cuando el modelo considera que el formato alternativo es más informativo. También puede ocurrir cuando el historial de datos tiene mucha variabilidad y Copilot prefiere expresar incertidumbre como rango.

**Solución:**
1. Envía el siguiente prompt de corrección en la misma conversación:
   ```
   Por favor, reformatea tu respuesta anterior como una tabla Markdown con 
   exactamente 4 columnas: Período, Escenario_Optimista, Escenario_Conservador, 
   Escenario_Crítico. Usa un único valor numérico entero por celda, 
   sin rangos ni texto adicional dentro de la tabla.
   ```
2. Si el problema persiste, extrae los valores manualmente del texto narrativo y constrúyelos en la tabla de Excel directamente.
3. Como última alternativa, calcula los escenarios manualmente en Excel usando las reglas definidas:
   - Optimista: `=PROMEDIO(últimos 3 meses) * 1.12`
   - Conservador: `=PROMEDIO(últimos 3 meses) * 1.00`
   - Crítico: `=PROMEDIO(últimos 3 meses) * 0.85`

---

### Problema 2 — El gráfico de líneas muestra los períodos como números (1, 2, 3) en lugar de M19, M20, M21

**Síntoma:** Al insertar el gráfico, el eje horizontal muestra los valores 1, 2, 3 en lugar de las etiquetas M19, M20, M21. Las series pueden aparecer correctamente, pero la lectura del eje X es confusa.

**Causa:** Excel interpreta la columna A (Período) como una serie de datos adicional en lugar de como etiquetas del eje horizontal, especialmente cuando el rango seleccionado incluye la columna A con texto.

**Solución:**
1. Haz clic derecho sobre el gráfico → `Seleccionar datos`.
2. En el panel izquierdo ("Entradas de leyenda / Series"), verifica que `Período` **no** aparezca como una serie. Si aparece, selecciónala y haz clic en **Quitar**.
3. En el panel derecho ("Etiquetas del eje horizontal"), haz clic en **Editar**.
4. En el campo "Rango de etiquetas del eje", selecciona el rango **A2:A4** (las celdas con M19, M20, M21).
5. Haz clic en **Aceptar** dos veces. El eje X ahora debe mostrar M19, M20, M21.

---

## Limpieza

Al finalizar la práctica, realiza las siguientes acciones de cierre:

1. **Guarda el archivo final** con `Ctrl + S`. Verifica que el nombre sea exactamente `Practica27_Comparacion_Excel_IA.xlsx`.

2. **Cierra la sesión de Copilot** si utilizaste una cuenta institucional, para proteger la privacidad de los datos compartidos en el prompt.

3. **Conserva el archivo** `Practica21_Pronostico_Seleccionado.xlsx` sin modificaciones — fue usado como referencia y no debe alterarse.

4. **Copia de respaldo:** Si el instructor lo solicita, guarda una copia del archivo en la carpeta compartida del curso con la nomenclatura: `[TuNombre]_Practica27_Comparacion_Excel_IA.xlsx`.

5. **No elimines** la hoja `Datos_Prompt` — el prompt construido en esta práctica puede ser reutilizado en prácticas posteriores con modificaciones menores.

---

## Resumen

### Lo que construiste en esta práctica

En 11 minutos ejecutaste un ciclo completo de análisis comparativo entre herramientas:

1. **Preparaste un prompt estructurado** con historial anonimizado y solicitud explícita de tres escenarios cuantitativos, aplicando las buenas prácticas de la Lección 6.1 (contextualizar, especificar formato, verificar coherencia).

2. **Obtuviste valores numéricos de Copilot** y los integraste en una tabla comparativa de 6 columnas en Excel, calculando diferencias porcentuales que evidencian la convergencia o divergencia entre el modelo estadístico y la estimación de IA.

3. **Visualizaste las cuatro series simultáneamente** en un gráfico de líneas, herramienta fundamental para comunicar la incertidumbre de demanda a equipos de operaciones e inventario.

4. **Documentaste criterio crítico** respondiendo 5 preguntas de análisis que te obligan a posicionarte sobre cuándo cada herramienta es más confiable — una habilidad directamente aplicable en reuniones de S&OP y planificación de reabastecimiento.

5. **Sintetizaste el aprendizaje** en un mini-framework de decisión reutilizable que responde la pregunta operativa central: *¿cuándo uso Excel y cuándo uso Copilot?*

### Conexión con el contexto logístico

La diferencia porcentual entre el escenario conservador de IA y el pronóstico estadístico de Excel no es solo un número académico: **define el rango de incertidumbre que debe cubrir tu inventario de seguridad**. Si la diferencia es del 8%, tu política de stock de seguridad debe contemplar ese margen. Si la diferencia es del 0–2%, ambas herramientas convergen y el pronóstico base es confiable. Este es el tipo de decisión que conecta el análisis de datos con la operación real de picking y reabastecimiento.

### Recursos adicionales

| Recurso | Descripción | Acceso |
|---|---|---|
| Microsoft Learn: Copilot en Excel | Documentación oficial sobre uso de Copilot para análisis de datos | [learn.microsoft.com/copilot/excel](https://learn.microsoft.com/es-es/copilot/microsoft-365/use-copilot-in-excel) |
| Lección 6.1 del curso | Fundamentos de IA generativa y sus limitaciones | Material del curso |
| Lección 6.2 del curso | Construcción de prompts efectivos para análisis logístico | Material del curso (próxima lección) |
| McKinsey: Generative AI in Supply Chain | Casos de uso industriales de IA generativa en logística | [mckinsey.com](https://www.mckinsey.com/capabilities/operations/our-insights/generative-ai-in-supply-chain) |

> 🔗 **Próxima práctica:** La Práctica 28 utilizará el framework de decisión construido en esta práctica como punto de partida para diseñar un flujo de trabajo integrado Excel + Copilot para ciclos mensuales de planificación de demanda.

---

---

---LAB_START---
LAB_ID: 06-00-03
---MARKDOWN---
# Práctica 28 — Generación de Recomendaciones Automatizadas de Inventario

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 11 minutos |
| **Complejidad** | Alta |
| **Nivel Bloom** | Crear (Create) |
| **Módulo** | 6 — IA Generativa (Copilot) en Análisis de Demanda |
| **Práctica número** | 28 |
| **Posición en la secuencia** | Práctica culminante integradora (requiere Prácticas 18–27 completadas) |

---

## 2. Descripción General

Esta práctica es el **ejercicio integrador final del curso**. El participante diseñará y ejecutará el prompt más complejo del módulo en Microsoft Copilot para obtener recomendaciones específicas de reabastecimiento (ROP, EOQ simplificado, stock de seguridad y frecuencia de revisión) para los 8 SKUs trabajados a lo largo del curso. Las recomendaciones generadas por la IA serán documentadas en una plantilla Excel pre-diseñada, evaluadas críticamente con criterios de aplicabilidad operativa, y sintetizadas en un resumen ejecutivo de una página titulado **"Plan de Reabastecimiento Asistido por IA"**. Este entregable integra los análisis de Excel (Módulo 4), las visualizaciones de Power BI (Módulo 5) y las recomendaciones de Copilot (Módulo 6) en un único documento operativo coherente.

---

## 3. Objetivos de Aprendizaje

Al completar esta práctica, el participante será capaz de:

- [ ] **Diseñar y ejecutar** un prompt avanzado multi-parámetro en Copilot que incluya contexto de negocio, datos de demanda y solicitudes específicas de indicadores de inventario (ROP, EOQ, stock de seguridad).
- [ ] **Documentar y estructurar** las recomendaciones de IA en una plantilla Excel de evaluación, aplicando criterios de aplicabilidad operativa (directa / con ajustes / no aplicable).
- [ ] **Evaluar críticamente** las recomendaciones generadas por IA comparándolas con las mejores prácticas logísticas conocidas, identificando limitaciones y ajustes necesarios.
- [ ] **Crear** un resumen ejecutivo integrador de una página ("Plan de Reabastecimiento Asistido por IA") que consolide los hallazgos del curso en un entregable operativo accionable.

---

## 4. Prerrequisitos

### Conocimiento previo requerido
- Haber completado las Prácticas 18 a 27 del curso, especialmente:
  - **Práctica 21**: Cálculo de indicadores estadísticos (promedio, CV, desviación estándar) por SKU en Excel.
  - **Práctica 25**: Dashboard de demanda en Power BI con visualizaciones de tendencia y estacionalidad.
  - **Práctica 27**: Selección del modelo de pronóstico con menor error (MAE/MAPE) y pronóstico del próximo mes.
- Comprensión operativa de los conceptos: **punto de reorden (ROP)**, **stock de seguridad**, **lead time** y **costo de quiebre de stock**.
- Haber ejecutado al menos un prompt básico en Copilot (Prácticas 26–27).

### Acceso requerido
- Acceso activo a **Microsoft Copilot** (copilot.microsoft.com) o herramienta equivalente (ChatGPT, Gemini).
- **Microsoft Excel** (2016 o superior) con el archivo de resultados consolidados de las prácticas anteriores.
- **Microsoft Word o PowerPoint** para el entregable final (a elección del participante).
- Archivo de plantilla proporcionado por el instructor: `Plantilla_P28_Recomendaciones_IA.xlsx`.

---

## 5. Entorno del Laboratorio

### Hardware mínimo recomendado

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 4 GB | 8 GB |
| Almacenamiento libre | 2 GB | 4 GB |
| Resolución de pantalla | 1366×768 | 1920×1080 |
| Conexión a internet | 10 Mbps | 25 Mbps o superior |

### Software requerido

| Software | Versión | Uso en esta práctica |
|---|---|---|
| Microsoft Excel | 2016 o superior (Microsoft 365 recomendado) | Plantilla de evaluación de recomendaciones |
| Microsoft Copilot / ChatGPT / Gemini | Versión actualizada (web) | Generación de recomendaciones de inventario |
| Microsoft Word o PowerPoint | Cualquier versión reciente | Resumen ejecutivo integrador |
| Microsoft Edge / Google Chrome | Versión 2024 actualizada | Acceso a Copilot web |

### Configuración previa al laboratorio

Antes de iniciar, verifica que tienes los siguientes archivos disponibles en tu carpeta de trabajo:

```
📁 Carpeta de trabajo recomendada: C:\Logistica_Forecasting\Practica28\
├── Plantilla_P28_Recomendaciones_IA.xlsx   ← Proporcionado por el instructor
├── Resultados_P21_Estadisticos.xlsx         ← Generado en Práctica 21
├── Resultados_P27_Pronostico_Seleccionado.xlsx ← Generado en Práctica 27
└── (Opcional) Capturas_PowerBI_P25.png      ← Exportadas en Práctica 25
```

**Verificación rápida de acceso a Copilot** (30 segundos):
1. Abre Microsoft Edge o Chrome.
2. Navega a `https://copilot.microsoft.com`.
3. Confirma que puedes iniciar una nueva conversación (botón "Nuevo tema" o "New topic").
4. Si no tienes acceso, usa `https://chat.openai.com` (ChatGPT) o `https://gemini.google.com` como alternativa.

---

## 6. Procedimiento Paso a Paso

> ⏱️ **Gestión del tiempo**: Esta práctica tiene 11 minutos. Distribución sugerida:
> - **Paso 1** (Preparación del contexto): 2 minutos
> - **Paso 2** (Construcción y ejecución del prompt): 3 minutos
> - **Paso 3** (Documentación en plantilla Excel): 3 minutos
> - **Paso 4** (Construcción del resumen ejecutivo): 3 minutos

---

### Paso 1: Preparación del Contexto para el Prompt

**Objetivo**: Consolidar los datos clave de las prácticas anteriores que serán incorporados al prompt de Copilot, asegurando que la IA cuente con suficiente contexto para generar recomendaciones relevantes y específicas.

#### Instrucciones

1. Abre el archivo `Resultados_P27_Pronostico_Seleccionado.xlsx`. Localiza la hoja con el resumen de los 8 SKUs. Necesitarás los siguientes valores para cada SKU:
   - Demanda promedio mensual (calculada en Práctica 21).
   - Coeficiente de variación — CV% (calculado en Práctica 21).
   - Pronóstico del próximo mes (calculado en Práctica 27).
   - Modelo seleccionado y su MAPE% (calculado en Práctica 27).

2. En un bloc de notas o directamente en el navegador, prepara la **tabla de datos resumida** que pegarás en el prompt. Usa el siguiente formato de referencia (reemplaza los valores con los de tu análisis real):

```
SKU        | Dem_Prom_Mensual | CV%   | Pronostico_Mes_N+1 | Modelo_Seleccionado | MAPE%
-----------|-----------------|-------|--------------------|---------------------|------
SKU-001    | 320 unidades    | 18%   | 335 unidades       | Suavización Exp.    | 8.2%
SKU-002    | 85 unidades     | 42%   | 91 unidades        | Promedio Móvil 3M   | 15.7%
SKU-003    | 510 unidades    | 12%   | 528 unidades       | Regresión Lineal    | 6.1%
SKU-004    | 47 unidades     | 67%   | 52 unidades        | Promedio Móvil 3M   | 22.4%
SKU-005    | 230 unidades    | 21%   | 241 unidades       | Suavización Exp.    | 9.8%
SKU-006    | 178 unidades    | 15%   | 185 unidades       | Regresión Lineal    | 7.3%
SKU-007    | 62 unidades     | 55%   | 58 unidades        | Promedio Móvil 3M   | 19.1%
SKU-008    | 395 unidades    | 9%    | 410 unidades       | Regresión Lineal    | 5.4%
```

> 📝 **Nota importante**: Los valores de la tabla anterior son **de referencia**. Debes usar los valores reales que obtuviste en tus prácticas anteriores. Si no completaste las prácticas previas, solicita al instructor el archivo de checkpoint correspondiente.

3. Mantén esta tabla visible (en una ventana separada o impresa) mientras construyes el prompt en el Paso 2.

#### Resultado esperado
Una tabla de 8 SKUs con sus indicadores clave lista para ser incorporada al prompt de Copilot.

#### Verificación
Confirma que tienes valores para los 5 campos (Dem_Prom_Mensual, CV%, Pronóstico, Modelo y MAPE%) para los 8 SKUs antes de avanzar.

---

### Paso 2: Construcción y Ejecución del Prompt Avanzado en Copilot

**Objetivo**: Diseñar y ejecutar el prompt más complejo del módulo — un prompt multi-parámetro que incluya contexto de negocio, datos de demanda y solicitudes específicas de indicadores de inventario — y obtener recomendaciones estructuradas de Copilot.

#### Instrucciones

1. Abre una **nueva conversación** en Copilot (copilot.microsoft.com) o en tu herramienta de IA alternativa. Haz clic en "Nuevo tema" para asegurarte de que no hay contexto previo que interfiera.

2. Copia el siguiente **prompt base** en el cuadro de texto de Copilot. **Antes de enviarlo**, reemplaza la sección `[TABLA DE DATOS]` con la tabla real que preparaste en el Paso 1:

```
Actúa como un consultor experto en gestión de inventarios y logística de almacén.
Voy a proporcionarte el contexto de negocio y datos de demanda histórica de 8 SKUs
para que generes recomendaciones específicas de reabastecimiento.

## CONTEXTO DEL NEGOCIO:
- Tipo de operación: Centro de distribución de consumo masivo (B2B y B2C)
- Lead time del proveedor: 2 semanas (14 días calendario)
- Costo de almacenamiento: 20% anual sobre el valor del inventario
- Costo de quiebre de stock: 35% del valor del pedido perdido
- Ciclo de revisión actual: mensual (se evalúa si debe cambiar)
- Nivel de servicio objetivo: 95% (Z = 1.65)
- Días hábiles por mes: 22 días

## DATOS DE DEMANDA (últimos 24 meses, análisis completado):
[TABLA DE DATOS — pega aquí tu tabla del Paso 1]

## SOLICITUD:
Para cada uno de los 8 SKUs, genera una tabla con las siguientes recomendaciones:
1. **ROP (Punto de Reorden)**: en unidades. Usa la fórmula: 
   ROP = Demanda_Diaria_Promedio × Lead_Time_Días + Stock_Seguridad
2. **Stock de Seguridad**: en unidades. Usa la fórmula:
   SS = Z × Desviación_Estándar_Diaria × √Lead_Time_Días
   (Calcula la desviación estándar diaria a partir del CV% y la demanda mensual promedio)
3. **EOQ simplificado**: en unidades. Usa la fórmula EOQ básica:
   EOQ = √(2 × Demanda_Anual × Costo_Pedido / Costo_Almacenamiento_Unitario)
   Asume: Costo por orden de compra = $150 USD; Valor unitario promedio = $12 USD
4. **Frecuencia de revisión recomendada**: semanal, quincenal o mensual,
   justificada por el CV% del producto.
5. **Prioridad de atención**: Alta / Media / Baja, basada en el CV% y el MAPE%.

Presenta los resultados en una tabla clara con columnas:
SKU | ROP_Recomendado | Stock_Seguridad | EOQ_Sugerido | Frecuencia_Revisión | Prioridad | Razonamiento_Breve

Después de la tabla, agrega un párrafo de máximo 5 líneas con las 3 recomendaciones
operativas más importantes para el gerente de inventarios.
```

3. **Envía el prompt** y espera la respuesta completa de Copilot (puede tardar 15–30 segundos).

4. **Lee la respuesta completa** antes de proceder. Si Copilot no generó la tabla en el formato solicitado, envía el siguiente **prompt de refinamiento**:

```
Por favor, reformatea la respuesta anterior presentando los resultados
exclusivamente en una tabla con las 7 columnas que solicité:
SKU | ROP_Recomendado | Stock_Seguridad | EOQ_Sugerido | 
Frecuencia_Revisión | Prioridad | Razonamiento_Breve
Mantén los valores numéricos calculados.
```

5. Una vez obtenida la tabla completa, **toma una captura de pantalla** de la respuesta (usa `Windows + Shift + S` en Windows o `Cmd + Shift + 4` en Mac). Guárdala como `Captura_Copilot_P28.png` en tu carpeta de trabajo.

#### Resultado esperado
Una tabla de 8 filas con recomendaciones de ROP, stock de seguridad, EOQ, frecuencia de revisión y prioridad para cada SKU, más un párrafo con las 3 recomendaciones operativas principales.

**Ejemplo de fragmento de respuesta esperada de Copilot**:

```
| SKU     | ROP_Recomendado | Stock_Seguridad | EOQ_Sugerido | Frecuencia_Revisión | Prioridad | Razonamiento_Breve                          |
|---------|----------------|-----------------|--------------|---------------------|-----------|---------------------------------------------|
| SKU-001 | 387 unidades   | 97 unidades     | 284 unidades | Quincenal           | Media     | CV moderado (18%), MAPE bajo, demanda estable|
| SKU-002 | 118 unidades   | 53 unidades     | 146 unidades | Semanal             | Alta      | Alta variabilidad (CV 42%), MAPE elevado     |
| ...     | ...            | ...             | ...          | ...                 | ...       | ...                                          |
```

#### Verificación
- La tabla contiene exactamente 8 filas (una por SKU).
- Todos los valores de ROP y Stock de Seguridad son positivos y expresados en unidades.
- La columna "Frecuencia_Revisión" contiene únicamente: Semanal, Quincenal o Mensual.
- Existe el párrafo de recomendaciones operativas al final de la respuesta.

---

### Paso 3: Documentación y Evaluación Crítica en la Plantilla Excel

**Objetivo**: Trasladar las recomendaciones de Copilot a la plantilla Excel estructurada y aplicar una evaluación crítica de aplicabilidad operativa para cada recomendación, identificando limitaciones de la IA.

#### Instrucciones

1. Abre el archivo `Plantilla_P28_Recomendaciones_IA.xlsx`. Localiza la hoja llamada **"Evaluación_Recomendaciones"**. Verás las siguientes columnas pre-configuradas:

```
A: Producto (SKU)
B: Demanda_Promedio_Mensual
C: Variabilidad_(CV%)
D: Pronóstico_Próximo_Mes
E: ROP_Recomendado_IA
F: Stock_Seguridad_IA
G: EOQ_Sugerido_IA
H: Frecuencia_Revisión_IA
I: Prioridad_IA
J: Aplicabilidad   ← "Directa" / "Con ajustes" / "No aplicable"
K: Observaciones_Críticas  ← Máximo 2 líneas de justificación
```

2. Completa las columnas **A a D** con los datos de tus prácticas anteriores (los mismos que usaste en el prompt del Paso 1).

3. Completa las columnas **E a I** transcribiendo los valores de la tabla generada por Copilot en el Paso 2.

4. Para cada SKU, completa la columna **J (Aplicabilidad)** seleccionando una de las tres opciones. Usa los siguientes criterios de evaluación:

| Criterio | Aplicable Directamente | Aplicable con Ajustes | No Aplicable |
|---|---|---|---|
| **Precisión del ROP** | ROP ≤ 1.2× tu cálculo manual | ROP difiere 20–50% de tu estimación | ROP difiere >50% o es incoherente |
| **Stock de Seguridad** | Coherente con el nivel de servicio 95% | Subestimado o sobreestimado moderadamente | Ignora la variabilidad real del SKU |
| **EOQ** | Supuestos de costo razonables para tu contexto | Supuestos de costo no representativos | Fórmula incorrectamente aplicada |
| **Frecuencia de revisión** | Alineada con el CV% del producto | Podría ajustarse según capacidad operativa | Contradice el comportamiento de demanda |

5. Para la columna **K (Observaciones_Críticas)**, escribe una justificación breve de máximo 2 líneas. Ejemplos orientativos:

```
Ejemplos de observaciones críticas:

✅ Aplicable directamente:
"ROP calculado es consistente con lead time de 14 días y CV bajo (9%).
Stock de seguridad adecuado para nivel de servicio 95%."

⚠️ Aplicable con ajustes:
"EOQ asume costo de pedido fijo de $150 USD; en nuestra operación el costo
real varía entre $80–$200 según proveedor. Requiere validación con compras."

❌ No aplicable:
"Copilot no consideró que SKU-004 tiene demanda lumpy (pedidos esporádicos
grandes). ROP continuo no aplica; se recomienda modelo (s,S) o revisión periódica."
```

6. Al finalizar la tabla, añade una fila de **totales/resumen** al pie de la hoja con la siguiente fórmula de conteo de aplicabilidad:

```excel
=COUNTIF(J2:J9,"Directa")          → Celda M2: "Recomendaciones aplicables directamente"
=COUNTIF(J2:J9,"Con ajustes")      → Celda M3: "Recomendaciones que requieren ajuste"
=COUNTIF(J2:J9,"No aplicable")     → Celda M4: "Recomendaciones no aplicables"
```

7. **Guarda el archivo** como `P28_Evaluacion_Completada_[TuNombre].xlsx`.

#### Resultado esperado
Una plantilla Excel con las 8 filas completadas, incluyendo las columnas A–I con datos y recomendaciones, y las columnas J–K con la evaluación crítica de cada recomendación. El conteo de aplicabilidad debe mostrar al menos 3 categorías distintas, evidenciando análisis crítico genuino.

#### Verificación
- Todas las celdas de las columnas A–K están completadas (sin celdas vacías).
- La columna J contiene únicamente los tres valores válidos: "Directa", "Con ajustes" o "No aplicable".
- Las observaciones de la columna K tienen un máximo de 2 líneas cada una.
- El conteo en M2:M4 suma exactamente 8.

---

### Paso 4: Construcción del Resumen Ejecutivo Integrador

**Objetivo**: Crear el entregable final del curso — un documento de una página titulado "Plan de Reabastecimiento Asistido por IA" — que integre los hallazgos de Excel, Power BI y Copilot en un formato comunicable a la gerencia operativa.

#### Instrucciones

1. Abre **Microsoft Word** o **Microsoft PowerPoint** (a tu elección). Crea un nuevo documento en blanco.

2. Usa la siguiente **estructura obligatoria** para el resumen ejecutivo. El documento debe caber en **una sola página** (Word) o **una sola diapositiva** (PowerPoint):

---

**ESTRUCTURA DEL RESUMEN EJECUTIVO**

```
┌─────────────────────────────────────────────────────────────────────────┐
│          PLAN DE REABASTECIMIENTO ASISTIDO POR IA                       │
│          [Nombre del participante] | [Fecha] | [Nombre del curso]       │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  SECCIÓN 1: HALLAZGOS CLAVE DEL ANÁLISIS DE DEMANDA (3 hallazgos)      │
│  ─────────────────────────────────────────────────────────────────────  │
│  1. [Hallazgo sobre tendencia o estacionalidad identificada]            │
│  2. [Hallazgo sobre productos con alta vs. baja variabilidad (CV%)]     │
│  3. [Hallazgo sobre eventos extraordinarios o anomalías detectadas]     │
│                                                                         │
│  SECCIÓN 2: MODELO DE PRONÓSTICO SELECCIONADO                          │
│  ─────────────────────────────────────────────────────────────────────  │
│  Modelo: [Nombre del modelo] | MAPE promedio: [X%] | MAE promedio: [X] │
│  Justificación operativa: [1–2 líneas sobre por qué este modelo]       │
│                                                                         │
│  SECCIÓN 3: TOP 5 RECOMENDACIONES DE INVENTARIO (priorizadas)          │
│  ─────────────────────────────────────────────────────────────────────  │
│  1. [SKU de mayor prioridad]: ROP=[X], SS=[X], Frecuencia=[X]          │
│  2. [Segunda prioridad]: ...                                            │
│  3. [Tercera prioridad]: ...                                            │
│  4. [Cuarta prioridad]: ...                                             │
│  5. [Quinta prioridad]: ...                                             │
│                                                                         │
│  SECCIÓN 4: REFLEXIÓN SOBRE EL USO DE IA EN FORECASTING LOGÍSTICO      │
│  ─────────────────────────────────────────────────────────────────────  │
│  [3–5 líneas sobre valor y limitaciones de la IA generativa en este    │
│   contexto. Incluir: qué aportó Copilot, qué requirió ajuste humano,   │
│   y una recomendación sobre cómo integrarla en el flujo de trabajo]    │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

---

3. Para la **Sección 3 (Top 5 recomendaciones)**, ordena los SKUs de mayor a menor prioridad usando la columna "Prioridad_IA" de tu plantilla Excel:
   - Primero los marcados como **Alta** prioridad (ordenados por CV% descendente).
   - Luego los de prioridad **Media** (ordenados por MAPE% descendente).
   - Los de prioridad **Baja** al final.

4. Para la **Sección 4 (Reflexión)**, considera responder las siguientes preguntas guía en tu texto (3–5 líneas en total, no como lista):
   - ¿Qué tipo de análisis realizó Copilot en segundos que habría tomado mucho más tiempo manualmente?
   - ¿En qué casos las recomendaciones de la IA requirieron ajuste por conocimiento del contexto operativo real?
   - ¿Cómo recomendarías integrar Copilot en el proceso mensual de planificación de inventarios?

   **Ejemplo orientativo de reflexión**:
   > *"Copilot demostró ser altamente eficiente para calcular indicadores de inventario estándar (ROP, SS, EOQ) en múltiples SKUs simultáneamente, reduciendo el tiempo de análisis de horas a minutos. Sin embargo, sus recomendaciones asumieron supuestos de costo fijos que no reflejan la variabilidad real de nuestra operación, y no consideró restricciones de espacio de almacén ni comportamientos de demanda lumpy en productos de baja rotación. La IA generativa es más valiosa como punto de partida analítico que como decisor autónomo; su integración óptima ocurre cuando el analista logístico valida, ajusta y contextualiza las recomendaciones generadas antes de aplicarlas operativamente."*

5. **Guarda el documento** como:
   - Word: `P28_Resumen_Ejecutivo_[TuNombre].docx`
   - PowerPoint: `P28_Resumen_Ejecutivo_[TuNombre].pptx`

#### Resultado esperado
Un documento de una página (Word) o una diapositiva (PowerPoint) que contenga las 4 secciones completas: hallazgos clave, modelo seleccionado, top 5 recomendaciones priorizadas y reflexión sobre el uso de IA. El documento debe ser legible y comunicable a un gerente de operaciones sin conocimiento técnico profundo.

#### Verificación
- El documento tiene exactamente las 4 secciones en el orden indicado.
- La Sección 1 contiene exactamente 3 hallazgos.
- La Sección 3 contiene exactamente 5 recomendaciones priorizadas con valores de ROP, SS y frecuencia.
- La Sección 4 tiene entre 3 y 5 líneas y menciona al menos una limitación de la IA.
- El documento cabe en una sola página o diapositiva.

---

## 7. Validación y Prueba Final

Al finalizar los 4 pasos, realiza la siguiente **lista de verificación de entregables**:

| # | Entregable | Criterio de completitud | ¿Completado? |
|---|---|---|---|
| 1 | `Captura_Copilot_P28.png` | Muestra tabla de 8 SKUs con las 7 columnas solicitadas | ☐ |
| 2 | `P28_Evaluacion_Completada_[TuNombre].xlsx` | 8 filas completas + columnas J y K evaluadas + conteo M2:M4 | ☐ |
| 3 | `P28_Resumen_Ejecutivo_[TuNombre].docx/.pptx` | 4 secciones completas en una sola página/diapositiva | ☐ |
| 4 | **Calidad del prompt** | El prompt enviado a Copilot incluía: contexto de negocio, tabla de datos, lead time, costos y solicitud estructurada | ☐ |
| 5 | **Evaluación crítica** | La columna J contiene las 3 categorías distintas (no todos "Directa") | ☐ |

**Prueba de integración**: Verifica que los valores en la Sección 3 del resumen ejecutivo (Top 5 recomendaciones) son consistentes con los datos de la plantilla Excel. Los valores de ROP y Stock de Seguridad deben coincidir exactamente con las columnas E y F de tu archivo Excel.

---

## 8. Resolución de Problemas

### Problema 1: Copilot genera recomendaciones genéricas sin valores numéricos específicos

**Síntoma**: La respuesta de Copilot contiene frases como "el ROP depende de su contexto específico" o "se recomienda calcular el stock de seguridad según sus parámetros" en lugar de valores numéricos calculados.

**Causa**: El prompt no incluyó suficiente contexto cuantitativo, o la tabla de datos no fue correctamente pegada en el cuadro de texto de Copilot (puede haberse perdido el formato tabular).

**Solución**:
1. Verifica que la tabla de datos está correctamente pegada en el prompt — debe ser texto plano con separadores `|` o tabulaciones, no una imagen.
2. Agrega la siguiente instrucción explícita al inicio del prompt: *"Usa los datos numéricos que te proporciono para calcular valores específicos. No uses frases genéricas; proporciona números concretos para cada SKU."*
3. Si el problema persiste, divide el prompt en dos partes: primero envía el contexto del negocio y pide confirmación de comprensión; luego envía la tabla de datos y la solicitud de cálculos.
4. Como alternativa, envía el prompt en inglés — los modelos de IA generativa suelen tener mejor rendimiento numérico en inglés debido a su entrenamiento.

---

### Problema 2: La plantilla Excel no suma correctamente en las celdas M2:M4

**Síntoma**: Las fórmulas `COUNTIF` en las celdas M2:M4 devuelven 0 o un valor incorrecto aunque la columna J está completada.

**Causa**: Los valores en la columna J contienen espacios adicionales, mayúsculas inconsistentes o caracteres especiales invisibles que impiden que `COUNTIF` reconozca los valores como coincidencias exactas.

**Solución**:
1. Selecciona el rango J2:J9 y aplica `Datos → Texto en columnas → Delimitado → Finalizar` para limpiar espacios ocultos.
2. Verifica que los valores en la columna J usan exactamente las cadenas: `"Directa"`, `"Con ajustes"` o `"No aplicable"` (respeta mayúsculas y acentos).
3. Si usas Excel en español, confirma que la función es `CONTAR.SI` en lugar de `COUNTIF`:
   ```excel
   =CONTAR.SI(J2:J9,"Directa")
   =CONTAR.SI(J2:J9,"Con ajustes")
   =CONTAR.SI(J2:J9,"No aplicable")
   ```
4. Como verificación adicional, usa `=LARGO(J2)` para confirmar que la longitud del texto es la esperada (7, 10 y 12 caracteres respectivamente) y no contiene caracteres ocultos.

---

## 9. Limpieza del Entorno

Al finalizar la práctica, realiza las siguientes acciones para mantener tu entorno de trabajo organizado:

1. **Guarda y cierra** todos los archivos de Excel y Word/PowerPoint abiertos.
2. **Cierra la sesión de Copilot** haciendo clic en "Nuevo tema" para limpiar el contexto de la conversación (esto protege la confidencialidad de los datos que compartiste).
3. **Organiza tu carpeta de trabajo** confirmando que los 3 entregables están presentes:
   ```
   📁 C:\Logistica_Forecasting\Practica28\
   ├── Captura_Copilot_P28.png                    ✅
   ├── P28_Evaluacion_Completada_[TuNombre].xlsx   ✅
   └── P28_Resumen_Ejecutivo_[TuNombre].docx/.pptx ✅
   ```
4. **No elimines** los archivos de prácticas anteriores (`Resultados_P21_*.xlsx`, `Resultados_P27_*.xlsx`) — pueden ser necesarios para revisión por parte del instructor.
5. Si trabajaste en un equipo compartido o de laboratorio, cierra sesión en el navegador para proteger el acceso a tu cuenta de Microsoft/Google.

---

## 10. Resumen y Recursos Adicionales

### Síntesis de la Práctica

En esta práctica culminante integraste las tres capas del ecosistema de forecasting logístico construido a lo largo del curso:

| Capa | Herramienta | Contribución a esta práctica |
|---|---|---|
| **Cálculo estadístico** | Excel (Módulos 3–4) | Datos base: demanda promedio, CV%, MAE, MAPE |
| **Visualización** | Power BI (Módulo 5) | Contexto visual de tendencias y estacionalidad |
| **Recomendación automatizada** | Copilot (Módulo 6) | ROP, SS, EOQ y frecuencia de revisión por SKU |

Los tres aprendizajes fundamentales de esta práctica son:

1. **La calidad del prompt determina la calidad de la recomendación**: Un prompt con contexto de negocio, datos cuantitativos y solicitudes estructuradas produce recomendaciones operativamente útiles; un prompt vago produce respuestas genéricas inaplicables.

2. **La IA generativa es un acelerador, no un decisor**: Copilot puede calcular indicadores de inventario para 8 SKUs en segundos, pero no conoce las restricciones de espacio de tu almacén, los acuerdos comerciales con tus proveedores ni el comportamiento lumpy de tus productos de menor rotación. El juicio profesional del analista logístico sigue siendo indispensable.

3. **La evaluación crítica es parte del proceso, no opcional**: La columna de "Observaciones_Críticas" en la plantilla Excel no es un formalismo académico — representa la habilidad profesional más importante cuando se trabaja con IA: saber cuándo confiar, cuándo ajustar y cuándo rechazar una recomendación generada automáticamente.

### Conexión con el Ciclo de Forecasting Logístico

Esta práctica cierra el ciclo metodológico completo del curso:

```
Datos históricos → Depuración → Análisis estadístico → Modelado → 
Evaluación de error → Pronóstico → Recomendación de inventario → 
Entregable operativo accionable
```

Cada etapa de este ciclo ahora tiene una herramienta asociada que el participante ha practicado: Excel para el análisis cuantitativo, Power BI para la visualización y comunicación, y Copilot para la aceleración de la interpretación y la generación de recomendaciones.

### Recursos de Referencia

| Recurso | Descripción | Acceso |
|---|---|---|
| Microsoft Copilot — Documentación oficial | Guía de uso de Copilot en Microsoft 365 | [learn.microsoft.com/es-es/copilot](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview) |
| Fórmula EOQ — Referencia académica | Modelo de Wilson para cantidad económica de pedido | Ballou, R. H. — *Logística: Administración de la cadena de suministro* |
| Fórmula ROP con stock de seguridad | Cálculo del punto de reorden con nivel de servicio | Chopra & Meindl — *Supply Chain Management* (Cap. 11) |
| McKinsey — IA en Supply Chain | Casos de uso de IA generativa en operaciones logísticas | [mckinsey.com/capabilities/operations](https://www.mckinsey.com/capabilities/operations/our-insights/generative-ai-in-supply-chain) |
| Prompt Engineering Guide | Guía de diseño de prompts efectivos para LLMs | [promptingguide.ai/es](https://www.promptingguide.ai/es) |

---

> 🎓 **Felicitaciones**: Has completado la Práctica 28 y con ella el ciclo completo de prácticas del curso de Forecasting Logístico con IA. El documento "Plan de Reabastecimiento Asistido por IA" que has creado representa un entregable real que podrás adaptar y utilizar en tu contexto profesional.

---
LAB_END---
