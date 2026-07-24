#!/usr/bin/env python3
"""
Script para geocodificar direcciones usando OpenStreetMap (Nominatim)
Lee un archivo Excel, busca coordenadas para cada dirección y guarda los resultados
"""

import time
from openpyxl import load_workbook
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

def geocodificar_excel(archivo_entrada, archivo_salida=None):
    """
    Geocodifica direcciones desde un archivo Excel

    Args:
        archivo_entrada: Ruta del archivo Excel de entrada
        archivo_salida: Ruta del archivo Excel de salida (por defecto mismo nombre + _geocodificado)
    """

    if archivo_salida is None:
        archivo_salida = archivo_entrada.replace('.xlsx', '_geocodificado.xlsx')

    # Cargar workbook
    print(f"📂 Leyendo archivo: {archivo_entrada}")
    wb = load_workbook(archivo_entrada)
    ws = wb.active

    # Inicializar geocodificador (Nominatim de OpenStreetMap)
    geolocator = Nominatim(user_agent="geocodificador_excel")

    # Encontrar columna de direcciones (busca "Dirección" o "Address")
    direccion_col = None
    for col_idx, cell in enumerate(ws[1], 1):
        if cell.value and ('dirección' in str(cell.value).lower() or 'address' in str(cell.value).lower()):
            direccion_col = col_idx
            break

    if direccion_col is None:
        print("❌ No se encontró columna 'Dirección' o 'Address' en la primera fila")
        return

    # Agregar columnas de resultado si no existen
    max_col = ws.max_column
    lat_col = max_col + 1
    lon_col = max_col + 2

    ws.cell(row=1, column=lat_col, value="Latitud")
    ws.cell(row=1, column=lon_col, value="Longitud")

    # Geocodificar cada fila
    total_filas = ws.max_row - 1
    print(f"🌍 Geocodificando {total_filas} direcciones...\n")

    for idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), 1):
        direccion = row[direccion_col - 1].value

        if not direccion:
            print(f"⏭️  Fila {idx}: Dirección vacía")
            continue

        try:
            print(f"⏳ Fila {idx}/{total_filas}: {direccion[:50]}...", end=" ", flush=True)

            # Geocodificar con timeout de 10 segundos
            location = geolocator.geocode(direccion, timeout=10)

            if location:
                ws.cell(row=idx + 1, column=lat_col, value=round(location.latitude, 6))
                ws.cell(row=idx + 1, column=lon_col, value=round(location.longitude, 6))
                print(f"✅ ({location.latitude:.6f}, {location.longitude:.6f})")
            else:
                print("⚠️  No encontrada")
                ws.cell(row=idx + 1, column=lat_col, value="No encontrada")
                ws.cell(row=idx + 1, column=lon_col, value="No encontrada")

            # Esperar 1 segundo entre requests para no sobrecargar Nominatim
            time.sleep(1)

        except GeocoderTimedOut:
            print("⏱️  Timeout")
            ws.cell(row=idx + 1, column=lat_col, value="Timeout")
            ws.cell(row=idx + 1, column=lon_col, value="Timeout")
            time.sleep(2)
        except GeocoderServiceError as e:
            print(f"❌ Error de servicio: {e}")
            ws.cell(row=idx + 1, column=lat_col, value="Error")
            ws.cell(row=idx + 1, column=lon_col, value="Error")
            time.sleep(2)
        except Exception as e:
            print(f"❌ Error: {e}")
            ws.cell(row=idx + 1, column=lat_col, value="Error")
            ws.cell(row=idx + 1, column=lon_col, value="Error")

    # Guardar archivo
    print(f"\n💾 Guardando archivo: {archivo_salida}")
    wb.save(archivo_salida)
    print("✨ ¡Geocodificación completada!")

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Uso: python geocodificar.py <archivo_excel>")
        print("Ejemplo: python geocodificar.py direcciones.xlsx")
        sys.exit(1)

    archivo = sys.argv[1]
    geocodificar_excel(archivo)
