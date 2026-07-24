# 📊 GEOCODIFICADOR EXCEL CON VBA

## ¿Qué hace?

Un archivo Excel que busca automáticamente **Latitud y Longitud** de direcciones usando OpenStreetMap (Nominatim) - ¡SIN necesidad de API key!

Simplemente:
1. ✏️ Agregas una dirección
2. 🔘 Presionas un botón (o Ctrl+Shift+G)
3. 📍 Excel busca automáticamente lat/long

---

## 📦 Archivos incluidos

- **Geocodificador.xlsx** - Archivo Excel listo para usar
- **CODIGO_VBA_GEOCODIFICADOR.bas** - Código de macros VBA
- **INSTRUCCIONES_VBA.md** - Este archivo

---

## 🚀 INSTALACIÓN PASO A PASO

### OPCIÓN 1: Importar el archivo .bas (Recomendado)

1. **Abre `Geocodificador.xlsx`** en Excel

2. **Habilita macros:**
   - Si Excel muestra una barra de advertencia amarilla
   - Haz clic en "Habilitar contenido"

3. **Abre el Editor de VBA:**
   - Presiona `ALT + F11` (Windows)
   - O en Mac: `Option + Fn + F11`

4. **Importa el código:**
   - En el panel izquierdo, verás un árbol de carpetas
   - Haz clic derecho en "VBAProject (Geocodificador.xlsx)"
   - Selecciona "Importar archivo..."
   - Abre `CODIGO_VBA_GEOCODIFICADOR.bas`
   - ✅ ¡Listo!

5. **Cierra el Editor de VBA:**
   - Presiona `ALT + F4` o cierra la ventana

---

### OPCIÓN 2: Copiar/Pegar el código manualmente

1. **Abre `Geocodificador.xlsx`**

2. **Abre el Editor de VBA:** `ALT + F11`

3. **Crea un nuevo módulo:**
   - Haz clic derecho en el árbol izquierdo
   - Selecciona "Insertar" → "Módulo"

4. **Copia el código:**
   - Abre `CODIGO_VBA_GEOCODIFICADOR.bas` en un editor de texto
   - Copia TODO el código
   - En el editor VBA, pégalo en el módulo nuevo

5. **Guarda el archivo:** `Ctrl + S`

---

## 📝 CÓMO USAR

### Uso básico:

1. **Abre `Geocodificador.xlsx`**

2. **En la hoja "Geocodificación":**
   - Columna A: Escribe tus direcciones
   - Ejemplo: `Avenida Gramado 1408, Puerto Varas, Los Lagos`

3. **Geocodifica:**
   - **Opción A:** Presiona `Ctrl + Shift + G`
   - **Opción B:** Busca el botón "Geocodificar" en la cinta
   - O en el Editor VBA (ALT+F11) presiona el botón ▶️ Play

4. **Resultado:**
   - Columna B: **Latitud** (ej: -41.3245)
   - Columna C: **Longitud** (ej: -72.4890)
   - Columna D: **Estado** (verde=✓ encontrada, rojo=✗ no encontrada)

---

## ⚙️ CONFIGURACIÓN

### Velocidad de procesamiento:

El código espera 1.5 segundos entre cada dirección para respetar los límites de Nominatim.

Para cambiar:
- Abre Editor VBA (ALT+F11)
- Busca: `Application.Wait Now + TimeValue("00:00:01.5")`
- Cambia `01.5` a lo que quieras (en segundos)

### Agregar más filas:

- El archivo viene con 15 filas (5 ejemplos + 10 vacías)
- Puedes agregar más direcciones en cualquier fila
- El código procesará automáticamente todas las que tengan contenido

---

## 🌍 REQUISITOS

✅ **Necesitas:**
- Excel 2007 o superior
- Conexión a Internet
- Macros habilitadas en Excel

---

## ⚠️ LIMITACIONES

- **Velocidad**: ~1-2 segundos por dirección
- **Precisión**: Depende de la calidad de la dirección
- **Disponibilidad**: Nominatim/OpenStreetMap puede tener picos de carga

### Tips para mejor precisión:

✅ Usa direcciones completas:
- `Avenida Gramado 1408, Puerto Varas, Los Lagos, Chile`

❌ Evita direcciones incompletas:
- `Avenida Gramado`

---

## 🔧 SOLUCIÓN DE PROBLEMAS

### "Excel no geocodifica nada"

1. Verifica que hayas habilitado macros
2. Verifica que tengas conexión a Internet
3. Prueba con una dirección más completa
4. Revisa el estado de Nominatim: https://nominatim.org/status/

### "Error: Usuario no autorizado"

- Es un error de conexión con Nominatim
- Espera unos minutos y vuelve a intentar
- O verifica tu conexión a Internet

### "Dirección no encontrada"

- Prueba con una dirección más completa
- Incluye ciudad y país
- Verifica que la dirección exista
- Nominatim es menos preciso con direcciones rurales

---

## 📚 REFERENCIA DE CÓDIGO

### Función principal:
```vba
Geocodificar()
```
Procesa todas las direcciones de la hoja activa.

### Función para obtener coordenadas:
```vba
ObtenerCoordenadas(direccion, lat, lon)
```
Consulta Nominatim y retorna True/False.

---

## 🔒 SEGURIDAD

- ✅ Código VBA es de código abierto
- ✅ Solo consulta Nominatim/OpenStreetMap
- ✅ NO envía datos a servidores externos (excepto a Nominatim)
- ✅ Respeta los términos de servicio de OpenStreetMap

---

## 📞 SOPORTE

Si tienes problemas:

1. Verifica que tengas habilitadas las macros
2. Prueba con una dirección simple
3. Comprueba tu conexión a Internet
4. Consulta el estado de Nominatim

---

## 🎯 PRÓXIMOS PASOS

Ahora puedes:
- ✏️ Agregar tus propias direcciones
- 🔘 Geocodificar con un clic
- 📊 Exportar los resultados
- 🗺️ Usar las coordenadas en mapas

¡Que disfrutes! 🎉

---

**Última actualización:** 2024
**Versión:** 1.0
**Licencia:** Código abierto
