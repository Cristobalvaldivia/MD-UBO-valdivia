# Geocodificación de Direcciones - Excel con OpenStreetMap

Sistema de geocodificación basado en código abierto que calcula latitud y longitud a partir de direcciones sin necesidad de APIs de pago.

## 📋 Características

- **Usar Nominatim (OpenStreetMap)**: Servicio gratuito de código abierto
- **Sin dependencia de Google Maps**: No requiere claves API
- **Automatizado**: Procesa múltiples direcciones en un solo paso
- **Excel compatible**: Trabaja directamente con archivos .xlsx

## 🚀 Instalación

```bash
pip install -r requirements.txt
```

## 📝 Uso

### 1. Preparar el archivo Excel

Tu archivo Excel debe tener una columna llamada **"Dirección"** (o "Address" en inglés).

Estructura recomendada:
| Dirección | Ciudad | País |
|-----------|--------|------|
| Calle 1, Santiago | Santiago | Chile |
| Paseo de la Reforma 505 | Ciudad de México | México |

### 2. Ejecutar geocodificación

```bash
python geocodificar.py direcciones.xlsx
```

Se creará un nuevo archivo: `direcciones_geocodificado.xlsx` con las columnas:
- **Latitud**: Coordenada de latitud (decimal)
- **Longitud**: Coordenada de longitud (decimal)

### 3. Resultado

El archivo de salida tendrá:
| Dirección | Ciudad | País | Latitud | Longitud |
|-----------|--------|------|---------|----------|
| Calle 1, Santiago | Santiago | Chile | -33.437346 | -70.673676 |

## ⚙️ Configuración

- **Espera entre requests**: 1 segundo (para no sobrecargar Nominatim)
- **Timeout**: 10 segundos por dirección
- **Precisión**: 6 decimales

## ⚠️ Limitaciones

- **Velocidad**: Nominatim tiene límites de rate limiting (~1 dirección/segundo)
- **Precisión variable**: Depende de la calidad de la dirección
- **Direcciones incompletas**: Pueden no encontrarse

## 💡 Tips

- Usa direcciones lo más completas posible (incluir ciudad/país)
- Las direcciones en español funcionan mejor para Latinoamérica
- Si una dirección no se encuentra, intenta ser más específico

## 📦 Dependencias

- **openpyxl**: Para leer/escribir archivos Excel
- **geopy**: Para geocodificación con Nominatim

## 🔗 Referencias

- [Nominatim OpenStreetMap](https://nominatim.org/)
- [Documentación de Geopy](https://geopy.readthedocs.io/)

---

**Nota**: Este proyecto utiliza el servicio de Nominatim/OpenStreetMap respetando sus términos de servicio.
