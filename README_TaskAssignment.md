# 📋 Módulo de Asignación de Tareas

## Descripción

Módulo para crear tablas de asignación de tareas de diseño en Google Slides. Permite distribuir temas entre diseñadores según su carga de trabajo (Stills y Conceptuales).

## Archivos del Módulo

- **TaskAssignmentModule.js** - Lógica principal del módulo
- **sidebar_v27.html** - Interfaz de usuario actualizada con nueva sección

## Características

### ✨ Funcionalidades

1. **Gestión de Diseñadores**
   - Agregar diseñadores con nombre
   - Definir cantidad de piezas: Stills y Conceptuales
   - Ver total de piezas por diseñador
   - Eliminar diseñadores de la lista

2. **Gestión de Temas**
   - Agregar temas con numeración correlativa (1, 2, 3...)
   - Soporte para sub-categorías (1a, 1b, 1c, 2, 3a, 3b, etc.)
   - Ordenamiento automático natural
   - Generación automática de temas según total de piezas
   - Validación de formato
   - Eliminar temas individuales o limpiar todos

3. **Generación de Tabla**
   - Crea tabla automática en nuevo slide
   - Primera columna: nombres de diseñadores con cantidades (S:x C:y)
   - Primera fila: temas numerados
   - Distribución correlativa de temas entre diseñadores
   - Marcas visuales (✓) para asignaciones

## Uso

### Paso 1: Agregar Diseñadores

1. En la sección **"📋 Asignaciones de Tareas"**
2. Ingresar nombre del diseñador
3. Ingresar cantidad de Stills
4. Ingresar cantidad de Conceptuales
5. Click en **"+ Agregar Diseñador"** o presionar Enter

**Ejemplo:**
```
Nombre: Ana García
Stills: 3
Conceptuales: 2
Total: 5 piezas
```

### Paso 2: Agregar Temas

**Opción A: Manual**
1. Ingresar tema en formato: `1`, `1a`, `2`, `3b`, etc.
2. Click en **"+ Agregar"** o presionar Enter
3. Los temas se ordenan automáticamente

**Opción B: Automático**
1. Click en **"Auto: 1, 2, 3..."**
2. Genera temas numerados (1, 2, 3...) según el total de piezas

**Formatos válidos de temas:**
- `1`, `2`, `3`, `10`, `25` (solo números)
- `1a`, `1b`, `2a`, `3c` (número + letra)

**Formatos NO válidos:**
- `a1` (letra primero)
- `1.5` (decimales)
- `1-a` (con guiones)
- `tema1` (texto adicional)

### Paso 3: Generar Tabla

1. Click en **"🎨 Generar Tabla de Asignación"**
2. La tabla se crea en un nuevo slide al final de la presentación
3. Los temas se asignan correlativamente a cada diseñador

## Ejemplo Completo

### Entrada:

**Diseñadores:**
- Ana García: 3 Stills, 2 Conceptuales = 5 piezas
- Carlos López: 2 Stills, 1 Conceptual = 3 piezas
- María Torres: 1 Still, 2 Conceptuales = 3 piezas

**Temas:** 1a, 1b, 2, 3a, 3b, 4, 5, 6, 7, 8, 9, 10, 11

### Salida (Tabla generada):

```
┌─────────────────┬────┬────┬───┬────┬────┬───┬───┬───┬───┬────┬────┬────┬────┐
│                 │ 1a │ 1b │ 2 │ 3a │ 3b │ 4 │ 5 │ 6 │ 7 │ 8  │ 9  │ 10 │ 11 │
├─────────────────┼────┼────┼───┼────┼────┼───┼───┼───┼───┼────┼────┼────┼────┤
│ Ana García      │ ✓  │ ✓  │ ✓ │ ✓  │ ✓  │   │   │   │   │    │    │    │    │
│ (S:3 C:2)       │    │    │   │    │    │   │   │   │   │    │    │    │    │
├─────────────────┼────┼────┼───┼────┼────┼───┼───┼───┼───┼────┼────┼────┼────┤
│ Carlos López    │    │    │   │    │    │ ✓ │ ✓ │ ✓ │   │    │    │    │    │
│ (S:2 C:1)       │    │    │   │    │    │   │   │   │   │    │    │    │    │
├─────────────────┼────┼────┼───┼────┼────┼───┼───┼───┼───┼────┼────┼────┼────┤
│ María Torres    │    │    │   │    │    │   │   │   │ ✓ │ ✓  │ ✓  │    │    │
│ (S:1 C:2)       │    │    │   │    │    │   │   │   │   │    │    │    │    │
└─────────────────┴────┴────┴───┴────┴────┴───┴───┴───┴───┴────┴────┴────┴────┘
```

## Validaciones

### Diseñadores
- ❌ Nombre vacío
- ❌ 0 piezas (sin Stills ni Conceptuales)
- ✅ Al menos 1 pieza (Still o Conceptual)

### Temas
- ❌ Formato inválido (no cumple patrón `\d+[a-z]?`)
- ❌ Temas duplicados
- ✅ Ordenamiento automático natural (1, 1a, 1b, 2, 2a, 3...)

### Generación
- ⚠️ Advertencia si hay más temas que piezas totales
- ❌ Error si no hay diseñadores
- ❌ Error si no hay temas

## Funciones Principales

### JavaScript (TaskAssignmentModule.js)

```javascript
generateTaskAssignmentTable(assignmentData)
```
- **Parámetros:** `{designers: Array, topics: Array}`
- **Retorna:** `{success: boolean, log: string, slideId: string}`

```javascript
createAssignmentMatrix(designers, topics)
```
- Crea matriz de asignación correlativa
- **Retorna:** Array 2D de booleanos

```javascript
sortTopics(topics)
```
- Ordena temas en orden natural
- **Retorna:** Array ordenado

### HTML/JavaScript (sidebar_v27.html)

```javascript
addDesigner()           // Agregar diseñador a la lista
removeDesigner(index)   // Eliminar diseñador
addTopic()              // Agregar tema
removeTopic(index)      // Eliminar tema
autoNumberTopics()      // Generar temas automáticos
clearTopics()           // Limpiar todos los temas
generateAssignmentTable() // Generar tabla en Google Slides
```

## Atajos de Teclado

- **Enter** en campos de diseñador → Agregar diseñador
- **Enter** en campo de tema → Agregar tema

## Notas Técnicas

### Tamaño de Tabla
- Ancho: 9 pulgadas
- Alto: 0.4 pulgadas × número de filas
- Posición: (0.5, 0.5) pulgadas desde esquina superior izquierda

### API Utilizada
- Google Slides Advanced API
- `Slides.Presentations.batchUpdate()`
- Request: `createTable`
- Request: `insertText`

### Distribución Correlativa
Los temas se asignan en orden secuencial:
1. Diseñador 1 recibe temas 1 hasta N1
2. Diseñador 2 recibe temas N1+1 hasta N1+N2
3. Y así sucesivamente...

## Mejoras Futuras

- [ ] Exportar/Importar asignaciones (JSON)
- [ ] Editar diseñadores después de agregarlos
- [ ] Drag & drop para reordenar temas
- [ ] Colores personalizados por diseñador
- [ ] Filtrar por tipo (Solo Stills, Solo Conceptuales)
- [ ] Estadísticas y balanceo de carga
- [ ] Guardar plantillas de equipos
