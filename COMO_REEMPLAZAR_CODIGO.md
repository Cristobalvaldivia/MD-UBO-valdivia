# 🔧 CÓMO REEMPLAZAR EL CÓDIGO VBA

Si tu código anterior da error, sigue estos pasos para reemplazarlo:

---

## OPCIÓN 1: Reemplazar el módulo (RECOMENDADO)

### Paso 1: Eliminar código viejo
1. **Abre Excel** - `Geocodificador.xlsx`
2. **Presiona ALT + F11** - Se abre el Editor de VBA
3. **En el panel izquierdo**, expande "Módulos"
4. **Haz clic derecho en "Módulo1"** (o el módulo con tu código)
5. Selecciona **"Eliminar"** 
6. Cuando pregunte "¿Deseas guardar los cambios?", haz clic en **"Sí"**

### Paso 2: Crear nuevo módulo
1. En el Editor de VBA, **haz clic derecho en "Módulos"**
2. Selecciona **"Insertar" → "Módulo"**
3. Se abre una ventana en blanco

### Paso 3: Copiar código nuevo
1. **Abre `CODIGO_VBA_SIMPLE.bas`** en un editor de texto (Notepad)
2. **Selecciona TODO el código** (Ctrl+A)
3. **Cópialo** (Ctrl+C)
4. **En el Editor de VBA**, pega el código (Ctrl+V)
5. **Cierra** el Editor (ALT+F4)
6. **Guarda Excel** (Ctrl+S)

### ¡Listo! Ahora presiona **Ctrl+Shift+G** para geocodificar

---

## OPCIÓN 2: Importar archivo .bas directamente

1. **Abre Excel** - `Geocodificador.xlsx`
2. **Presiona ALT + F11** - Editor de VBA
3. **Primero elimina el módulo viejo:**
   - Haz clic derecho en "Módulo1" 
   - Selecciona "Eliminar"
   - Confirma "Sí"

4. **Importa el nuevo módulo:**
   - Haz clic derecho en "Módulos"
   - Selecciona **"Importar archivo..."**
   - **Selecciona `CODIGO_VBA_SIMPLE.bas`**
   - ¡Listo!

---

## CUÁL CÓDIGO USAR

**Si tienes error:**
- ✅ Usa: `CODIGO_VBA_SIMPLE.bas` (MÁS SIMPLE, MENOS ERRORES)

**Si quieres mejor calidad:**
- 📦 Prueba: `CODIGO_VBA_MEJORADO.bas` (MÁS ROBUSTO)

**Diferencias:**

| Aspecto | Simple | Mejorado |
|---------|--------|----------|
| Complejidad | ⭐ Muy simple | ⭐⭐ Simple |
| Manejo de errores | Básico | Avanzado |
| Velocidad | Rápido | Normal |
| Precisión | Buena | Muy buena |

---

## 🧪 PRUEBA RÁPIDA

Después de cambiar el código:

1. **Ve a la hoja "Geocodificación"**
2. **En la columna A, agrega una dirección:**
   - Ejemplo: `Puerto Varas, Chile`
3. **Presiona Ctrl+Shift+G**
4. Si funciona:
   - ✅ Columnas B y C se llenarán con lat/long
   - ✅ Columna D mostrará "✓ Encontrada" en verde

Si sigue habiendo error:
- ✅ Asegúrate de habilitar macros en Excel
- ✅ Verificar que tengas conexión a Internet
- ✅ Prueba con una dirección más simple

---

## 🆘 SOLUCIÓN DE PROBLEMAS

### "Sigue sin funcionar"
- Cierra Excel completamente
- Abre nuevamente `Geocodificador.xlsx`
- Habilita macros si lo pregunta
- Intenta de nuevo

### "Error de Internet"
- Verifica tu conexión WiFi
- Intenta en unos minutos (Nominatim puede estar saturado)

### "No encuentra direcciones"
- Prueba con: `Santiago, Chile`
- O: `Buenos Aires, Argentina`
- Las direcciones deben ser válidas y completas

---

¿Funciona ahora? 🎉

Si aún hay problemas, los archivos están guardados en el repositorio para que puedas descargarlos nuevamente.
