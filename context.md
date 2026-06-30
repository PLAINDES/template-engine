# Módulo: `ai-variables`

## Objetivo

Autocompletar mediante IA (Gemini 2.5 Flash vía Vertex AI) las variables de un proyecto que no puedan obtenerse desde fuentes automáticas (SUNAT, ESCALE, APIs de ubigeo). El módulo soporta dos sistemas de plantillas con flujos independientes:

- **Sistema Editor:** variables estructuradas en `valoresProyecto` (objeto JSON en MinIO `object.json`)
- **Sistema DOCX:** variables planas detectadas del `.docx` almacenadas en `DocumentDraft.variables` (BD)

En ambos casos, Gemini recibe el contexto completo del proyecto para generar valores coherentes y técnicamente válidos para los campos vacíos.

---

## Responsabilidades

- Detectar qué variables están vacías en cada sistema y cuáles tienen fuente automática (excluidas de IA).
- Construir prompts con el contexto del proyecto relevante para cada sistema.
- Llamar a `GeminiService.generateJson()` y obtener el JSON con valores sugeridos.
- Garantizar que la IA nunca pise variables con valor existente.
- Persistir los resultados: `object.json` en MinIO (Editor) o `DocumentDraft` en BD (DOCX).
- Registrar auditoría por cada autocompletado realizado (sistema, modelo, campos llenados, confianza).

---

## Estructura de carpetas

```
ai-variables/
├── dto/
│   ├── autocompletar-editor.dto.ts       # { projectId: number }
│   ├── autocompletar-docx.dto.ts         # { projectId: number, draftId: number }
│   └── resultado-autocompletado.dto.ts   # respuesta unificada de ambos sistemas
├── entities/
│   └── ai-variable-change.entity.ts      # tabla ai_variable_changes — auditoría
├── utils/
│   ├── variables-excluidas-editor.ts     # paths de valoresProyecto excluidos de IA
│   └── build-prompt.ts                   # constructores de prompt para cada sistema
├── ai-variables.controller.ts
├── ai-variables.module.ts
└── ai-variables.service.ts
```

---

## Proveedor de IA

**Gemini 2.5 Flash** vía **Vertex AI** (`@google-cloud/vertexai`).

El servicio compartido `GeminiService` vive en `src/shared/intelligence/gemini.service.ts` y es el único punto de contacto con la API de Gemini. El módulo `GeminiModule` lo exporta para ser consumido por cualquier módulo que lo necesite.

### Métodos de `GeminiService` requeridos

```ts
// Ya implementado — para prompts simples de texto
async generateContent(prompt: string): Promise<string>

// Nuevo — para autocompletado: fuerza respuesta JSON válida
async generateJson(systemPrompt: string, userPrompt: string): Promise<Record<string, any>> {
  const result = await this.model.generateContent({
    contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
    systemInstruction: { parts: [{ text: systemPrompt }] },
    generationConfig: { responseMimeType: 'application/json' },
  });
  return JSON.parse(result.response.candidates[0].content.parts[0].text);
}
```

`responseMimeType: 'application/json'` fuerza a Gemini a responder siempre JSON válido, sin texto libre alrededor.

---

## Entidad de auditoría

### `AiVariableChange` (`ai_variable_changes`)

| Campo            | Tipo                  | Descripción                                                    |
|------------------|-----------------------|----------------------------------------------------------------|
| `id`             | number                | PK autoincremental                                             |
| `sistema`        | `'editor' \| 'docx'`  | Sistema de plantilla sobre el que se operó                     |
| `projectId`      | number                | FK hacia el proyecto                                           |
| `draftId`        | number?               | FK hacia `DocumentDraft` — solo en sistema DOCX                |
| `userId`         | string                | Cognito sub del usuario que solicitó el autocompletado         |
| `camposLlenados` | JSON                  | `Record<string, string>` — campo → valor generado              |
| `camposIgnorados`| JSON                  | Campos que ya tenían valor o tenían fuente automática          |
| `confianza`      | JSON                  | `Record<string, number>` — score por campo (0-1)               |
| `modelo`         | string                | Modelo utilizado, ej: `"gemini-2.5-flash"`                     |
| `createdAt`      | Date                  | Timestamp del autocompletado                                   |

---

## Variables excluidas — Sistema Editor

Las siguientes rutas de `valoresProyecto` **nunca se delegan a la IA** porque tienen fuente automática:

| Ruta excluida             | Fuente automática                                       |
|---------------------------|---------------------------------------------------------|
| `empresas`                | SUNAT vía Decolecta (RUC)                              |
| `institucionEducativa.*`  | ESCALE vía `schools.service` (sector Educación)        |
| `centroMedico.*`          | BD local vía `medicalCenters.service` (sector Salud)   |
| `ubicacion.*`             | Ubigeo resuelto desde `departmentId/provinceId/districtId` del proyecto |

El archivo `utils/variables-excluidas-editor.ts` exporta un `Set<string>` con estas rutas. Si se integra una nueva fuente automática, se actualiza solo este archivo.

---

## Variables excluidas — Sistema DOCX

En el sistema DOCX, las variables con fuente automática son las que aparecen en `DOCX_VARIABLE_MAP` (`document-upload/utils/docx-variable-mapping.ts`). Estas se mapean automáticamente desde `valoresProyecto` mediante `applyDocxVariableMapping()`.

La IA solo interviene sobre las variables del `DocumentDraft` que:
1. No están en `DOCX_VARIABLE_MAP` (sin fuente automática), **y**
2. Tienen valor vacío (`""`) en `DocumentDraft.variables`

---

## Prompts

### Sistema Editor

```
[SYSTEM]
Eres un redactor técnico especialista en proyectos de inversión pública del Perú (SNIP/Invierte.pe).
Tu tarea es completar los campos vacíos del objeto de variables de un proyecto.
Responde ÚNICAMENTE con un objeto JSON válido. Incluye un campo "confianza" por cada variable generada (número entre 0 y 1).
No inventes datos estructurados (códigos, RUC, coordenadas) — esos campos deben dejarse vacíos.

[USER]
## Contexto del proyecto
- Sector: {sectorNombre}
- Tipo de estudio: {tipoEstudioNombre}
- Naturaleza de intervención: {naturalezaNombre}
- Clasificación: {clasificacionNombre}
- Ubicación: {departamento} / {provincia} / {distrito}

## Variables ya completadas
{JSON.stringify(variablesConValor, null, 2)}

## Campos vacíos a completar
{JSON.stringify(camposVacíos)}  ← lista de nombres de campo

## Formato de respuesta esperado
{
  "campo1": "valor",
  "campo2": "valor",
  "confianza": { "campo1": 0.9, "campo2": 0.7 }
}
```

### Sistema DOCX

```
[SYSTEM]
Eres un redactor técnico especialista en proyectos de inversión pública del Perú (SNIP/Invierte.pe).
Tu tarea es completar variables de un documento técnico. Lee el contenido del documento y los datos
del proyecto para inferir valores coherentes con el contexto real.
Responde ÚNICAMENTE con un objeto JSON válido. Incluye un campo "confianza" por cada variable (0-1).

[USER]
## Datos del proyecto
- Sector: {sectorNombre}
- Tipo de estudio: {tipoEstudioNombre}
- Naturaleza: {naturalezaNombre}
- Ubicación: {departamento} / {provincia} / {distrito}

## Variables del proyecto ya resueltas
{JSON.stringify(valoresProyecto)}    ← object.json completo del proyecto

## Contenido del documento (texto extraído del .docx)
{textoDocumento}    ← texto completo o secciones relevantes del .docx

## Variables vacías a completar (formato [NOMBRE_VARIABLE])
{JSON.stringify(variablesVacíasDocx)}

## Formato de respuesta esperado
{
  "NOMBRE_VARIABLE_1": "valor",
  "NOMBRE_VARIABLE_2": "valor",
  "confianza": { "NOMBRE_VARIABLE_1": 0.85, "NOMBRE_VARIABLE_2": 0.7 }
}
```

> El texto del documento se obtiene llamando al microservicio Python (`DOCX_SERVICE_URL/extract-text`) pasando el `minioKey` del `DocumentDraft`. Si el microservicio no expone ese endpoint aún, se puede usar el contenido HTML ya generado que el frontend guarda en MinIO.

---

## DTOs

### `AutocompletarEditorDto`
```ts
{ projectId: number }
```

### `AutocompletarDocxDto`
```ts
{ projectId: number, draftId: number }
```

### `ResultadoAutocompletadoDto` (respuesta unificada)
```ts
{
  sistema: 'editor' | 'docx',
  camposLlenados: Record<string, string>,
  camposIgnorados: string[],
  confianza: Record<string, number>,
  modelo: string,
}
```

---

## Relación Controller → Service → Module

```
AiVariablesController
  └── AiVariablesService
        ├── GeminiService                        (Vertex AI — generateJson)
        ├── Repository<AiVariableChange>         (TypeORM — auditoría)
        ├── Repository<DocumentDraft>            (TypeORM — sistema DOCX)
        ├── ProyectsService                      (cargar proyecto con relaciones)
        └── HttpService / MinioClient            (leer/escribir object.json y .docx)

AiVariablesModule
  ├── imports:  TypeOrmModule.forFeature([AiVariableChange, DocumentDraft])
  │             GeminiModule
  │             ProjectsModule
  │             HttpModule (para llamar al microservicio Python)
  ├── controllers: [AiVariablesController]
  ├── providers:   [AiVariablesService]
  └── exports:     []
```

---

## Endpoints

| Método | Ruta                                       | Descripción                                                       |
|--------|--------------------------------------------|-------------------------------------------------------------------|
| `POST` | `/ai-variables/autocompletar/editor`       | Autocompletar `valoresProyecto` y actualizar `object.json`        |
| `POST` | `/ai-variables/autocompletar/docx`         | Autocompletar variables del `DocumentDraft` y guardar en BD       |
| `GET`  | `/ai-variables/historial/:projectId`       | Historial de autocompletados del proyecto (ambos sistemas)        |
| `GET`  | `/ai-variables/test`                       | Verificar conexión con Gemini (requiere Bearer Token)             |

Todos los endpoints requieren `Bearer Token` (Cognito).

---

## Flujos de datos

### Sistema Editor — `POST /ai-variables/autocompletar/editor`

```
Frontend
  → POST { projectId: 42 }
  → autocompletarEditor(projectId, userId)
      1. Cargar project entity        ← sector, estudio, naturaleza, departmentId, etc.
      2. Cargar object.json (MinIO)   ← valoresProyecto actual
      3. Identificar campos vacíos    ← valor === "" o undefined (recorre el objeto recursivamente)
      4. Filtrar excluidos            ← variables-excluidas-editor.ts
      5. Si no hay campos vacíos      → retornar { camposLlenados: {}, camposIgnorados: [...] }
      6. buildPrompt('editor', ...)   ← contexto del proyecto + campos vacíos
      7. geminiService.generateJson() ← Vertex AI, JSON mode
      8. Validar respuesta            ← solo aceptar keys que estaban en la lista de campos vacíos
      9. Merge en valoresProyecto     ← deep merge, solo campos vacíos
     10. Guardar object.json (MinIO)  ← versión actualizada
     11. INSERT AiVariableChange      ← sistema: 'editor', auditoría completa
  ← ResultadoAutocompletadoDto
```

### Sistema DOCX — `POST /ai-variables/autocompletar/docx`

```
Frontend
  → POST { projectId: 42, draftId: 7 }
  → autocompletarDocx(projectId, draftId, userId)
      1. Cargar DocumentDraft (BD)    ← variables (Record<string,string>), minioKey
      2. Cargar project entity        ← sector, estudio, naturaleza, ubicación
      3. Cargar object.json (MinIO)   ← valoresProyecto completo (contexto para IA)
      4. Aplicar DOCX_VARIABLE_MAP    ← llenar automáticamente lo que viene de valoresProyecto
      5. Identificar vacíos restantes ← variables no en el mapa Y con valor ""
      6. Si no hay vacíos restantes   → retornar { camposLlenados: {}, ... }
      7. Obtener texto del .docx      ← GET DOCX_SERVICE_URL/extract-text?key={minioKey}
      8. buildPrompt('docx', ...)     ← texto doc + valoresProyecto + variables vacías
      9. geminiService.generateJson() ← Vertex AI, JSON mode
     10. Validar respuesta            ← solo keys que estaban en lista de vacíos
     11. Merge en DocumentDraft.variables ← solo vacíos, nunca pisar existentes
     12. Guardar DocumentDraft (BD)   ← UPDATE variables
     13. INSERT AiVariableChange      ← sistema: 'docx', draftId, auditoría
  ← ResultadoAutocompletadoDto
```

### Historial — `GET /ai-variables/historial/:projectId`

```
Frontend
  → GET /ai-variables/historial/42
  → aiVariableChangeRepo.find({
      where: { projectId: 42 },
      order: { createdAt: 'DESC' }
    })
  ← [{ id, sistema, draftId, camposLlenados, confianza, modelo, createdAt }]
```

---

## Cambios necesarios en módulos existentes

### `GeminiService` (`src/shared/intelligence/gemini.service.ts`)

Agregar método `generateJson()` con JSON mode nativo de Vertex AI:

```ts
async generateJson(
  systemPrompt: string,
  userPrompt: string,
): Promise<Record<string, any>> {
  const result = await this.model.generateContent({
    contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
    systemInstruction: { parts: [{ text: systemPrompt }] },
    generationConfig: { responseMimeType: 'application/json' },
  });
  const text = result.response.candidates[0].content.parts[0].text;
  return JSON.parse(text);
}
```

### Microservicio Python (`DOCX_SERVICE_URL`)

Para el sistema DOCX, se necesita un endpoint que extraiga el texto del `.docx` almacenado en MinIO:

```
GET /extract-text?key={minioKey}
← { texto: "contenido completo del .docx como texto plano" }
```

Si este endpoint aún no existe en el microservicio Python, la alternativa es descargar el `.docx` desde MinIO y extraer el texto con la librería `docx` de Node.js ya instalada en el proyecto.

### `app.module.ts`

`AiVariablesModule` ya está registrado. ✓

---

## Dependencias importantes

- **`GeminiModule`** — exporta `GeminiService` con Vertex AI. Ya creado en `src/shared/intelligence/`.
- **`ProjectsModule`** — para cargar el proyecto con sus relaciones (sector, estudio, ubicación).
- **MinIO client** — para leer/escribir `object.json`. Usar el mismo cliente que usan otros servicios del proyecto.
- **`DocumentDraft` entity** — para leer y actualizar variables del sistema DOCX.
- **Microservicio Python** — para extracción de texto del `.docx` en el sistema DOCX.

---

## Consideraciones técnicas

- **Solo campos vacíos, sin excepciones:** después de que Gemini responde, el service filtra su respuesta para aceptar únicamente keys que estaban en la lista de `camposVacíos`. Cualquier campo que Gemini invente fuera de esa lista es descartado silenciosamente.

- **JSON mode nativo:** `responseMimeType: 'application/json'` garantiza que Gemini nunca devuelva texto libre, eliminando el riesgo de fallos en `JSON.parse`.

- **Contexto del `.docx` para sistema DOCX:** enviar el texto completo del documento es lo que diferencia el autocompletado genérico de uno que realmente entiende el contexto. Sin él, Gemini solo puede inferir desde los metadatos del proyecto. Con él, puede leer el diagnóstico, los antecedentes y la descripción del proyecto para generar variables coherentes con lo ya redactado.

- **Idempotencia:** ejecutar el autocompletado dos veces sobre el mismo proyecto/draft no rompe nada — la segunda ejecución no encontrará campos vacíos y retornará `camposLlenados: {}`. El historial registra ambas ejecuciones.

- **Modelo actual:** `gemini-2.5-flash` (configurado en `GeminiService`). Si se necesita mayor capacidad de razonamiento sobre documentos muy largos, cambiar a `gemini-2.5-pro` es una línea en `gemini.service.ts`.

- **Auditoría inmutable:** los registros de `AiVariableChange` no se modifican. Son la base para medir calidad comparando `camposLlenados` vs. el valor final al momento de generar el documento.

---

## Métricas de calidad (implementación futura)

| Métrica             | Cálculo                                                            |
|---------------------|--------------------------------------------------------------------|
| Tasa de aceptación  | Campos donde valor final == valor IA / total campos llenados       |
| Tasa de edición     | Distancia de edición promedio entre valor IA y valor final         |
| Cobertura           | Campos completados por IA / total campos vacíos disponibles        |
| Score de confianza  | Promedio del campo `confianza` devuelto por Gemini                 |

Para habilitarlas, el módulo que genera el documento debe leer el último `AiVariableChange` del proyecto y comparar contra los valores finales antes de renderizar.

---

## Ejemplos de solicitud y respuesta

### POST `/ai-variables/autocompletar/editor`

**Request:**
```json
{ "projectId": 42 }
```

**Response:**
```json
{
  "sistema": "editor",
  "camposLlenados": {
    "vial.tipoVia": "trocha carrozable",
    "vial.nombreVia": "Camino vecinal tramo Puente Capelo - Caserío El Molino",
    "vial.localidadOrigen": "Puente Capelo",
    "vial.localidadDestino": "Caserío El Molino"
  },
  "camposIgnorados": ["empresas", "ubicacion.departamentoNombre", "institucionEducativa.nombre"],
  "confianza": {
    "vial.tipoVia": 0.88,
    "vial.nombreVia": 0.74,
    "vial.localidadOrigen": 0.82,
    "vial.localidadDestino": 0.79
  },
  "modelo": "gemini-2.5-flash"
}
```

### POST `/ai-variables/autocompletar/docx`

**Request:**
```json
{ "projectId": 42, "draftId": 7 }
```

**Response:**
```json
{
  "sistema": "docx",
  "camposLlenados": {
    "DESCRIPCION_PROBLEMA": "La trocha carrozable presenta deterioro severo...",
    "OBJETIVO_PROYECTO": "Mejorar las condiciones de transitabilidad...",
    "BENEFICIARIOS_DIRECTOS": "1,240 habitantes del caserío El Molino"
  },
  "camposIgnorados": ["NOMBRE_PROYECTO", "NOMBRE_DEPARTAMENTO", "CODIGO_IE"],
  "confianza": {
    "DESCRIPCION_PROBLEMA": 0.81,
    "OBJETIVO_PROYECTO": 0.86,
    "BENEFICIARIOS_DIRECTOS": 0.70
  },
  "modelo": "gemini-2.5-flash"
}
```
