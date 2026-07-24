'===============================================
' CÓDIGO VBA PARA GEOCODIFICACIÓN EN EXCEL
' Basado en Nominatim (OpenStreetMap)
' Copiar este código en el Editor de VBA de Excel
'===============================================

' INSTRUCCIONES DE INSTALACIÓN:
' 1. Abre Geocodificador.xlsx
' 2. Presiona ALT+F11 para abrir el Editor de VBA
' 3. En el panel izquierdo, haz clic derecho en "VBAProject (Geocodificador.xlsx)"
' 4. Selecciona "Importar archivo..."
' 5. Selecciona este archivo (.bas)
' O copiar/pegar el código en un módulo nuevo

'===============================================
' MÓDULO PRINCIPAL - Geocodificación
'===============================================

Option Explicit

' Constantes
Const NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
Const USER_AGENT = "Geocodificador-Excel/1.0"

Sub Geocodificar()
    '
    ' Geocodificar - Procesa todas las direcciones en la hoja
    '
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim direccion As String
    Dim lat As Double
    Dim lon As Double
    Dim resultado As String
    Dim contador As Long

    ' Obtener la hoja activa
    Set ws = ThisWorkbook.Sheets("Geocodificación")

    ' Encontrar última fila con datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    contador = 0

    ' Procesar cada fila desde la 2 (saltando encabezado)
    For row = 2 To lastRow
        direccion = Trim(ws.Cells(row, 1).Value)

        ' Si hay dirección, geocodificar
        If Len(direccion) > 0 Then
            ws.Cells(row, 4).Value = "Procesando..."
            ws.Cells(row, 4).Interior.Color = RGB(255, 255, 0) ' Amarillo

            DoEvents ' Permitir que la pantalla se actualice

            ' Obtener coordenadas
            If ObtenerCoordenadas(direccion, lat, lon) Then
                ws.Cells(row, 2).Value = Format(lat, "0.000000")
                ws.Cells(row, 3).Value = Format(lon, "0.000000")
                ws.Cells(row, 4).Value = "✓ Encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(0, 176, 80) ' Verde
                contador = contador + 1
            Else
                ws.Cells(row, 2).Value = "N/A"
                ws.Cells(row, 3).Value = "N/A"
                ws.Cells(row, 4).Value = "✗ No encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(192, 0, 0) ' Rojo
            End If

            ' Esperar 1.5 segundos para no sobrecargar Nominatim
            Application.Wait Now + TimeValue("00:00:01.5")
        End If
    Next row

    MsgBox "✅ Geocodificación completada!" & vbCrLf & _
           "Direcciones procesadas: " & contador, vbInformation, "Éxito"

    Exit Sub
ErrorHandler:
    MsgBox "❌ Error: " & Err.Description, vbCritical, "Error"
End Sub

'===============================================
' FUNCIÓN AUXILIAR - Obtener Coordenadas
'===============================================

Function ObtenerCoordenadas(direccion As String, ByRef lat As Double, ByRef lon As Double) As Boolean
    '
    ' Consulta Nominatim para obtener coordenadas de una dirección
    ' Retorna True si tuvo éxito, False si no encuentra la dirección
    '
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim url As String
    Dim response As String
    Dim jsonObj As Object

    ' Crear objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Construir URL con la dirección
    url = NOMINATIM_URL & "?q=" & URLEncode(direccion) & "&format=json&limit=1"

    ' Realizar solicitud GET
    xmlHttp.Open "GET", url, False
    xmlHttp.SetRequestHeader "User-Agent", USER_AGENT
    xmlHttp.SetRequestHeader "Accept-Language", "es-ES,es;q=0.9"
    xmlHttp.Send

    response = xmlHttp.responseText

    ' Verificar que haya respuesta
    If Len(response) = 0 Or response = "[]" Then
        ObtenerCoordenadas = False
        Exit Function
    End If

    ' Parsear JSON (método simple)
    If InStr(response, """lat""") > 0 And InStr(response, """lon""") > 0 Then
        lat = CDbl(ExtractJSONValue(response, "lat"))
        lon = CDbl(ExtractJSONValue(response, "lon"))
        ObtenerCoordenadas = True
    Else
        ObtenerCoordenadas = False
    End If

    Exit Function
ErrorHandler:
    ObtenerCoordenadas = False
End Function

'===============================================
' FUNCIÓN AUXILIAR - Extraer Valor JSON
'===============================================

Function ExtractJSONValue(jsonText As String, keyName As String) As String
    '
    ' Extrae un valor de un JSON simple (búsqueda por texto)
    '
    Dim searchStr As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    searchStr = """" & keyName & """:"
    startPos = InStr(jsonText, searchStr)

    If startPos = 0 Then
        ExtractJSONValue = "0"
        Exit Function
    End If

    startPos = startPos + Len(searchStr)
    endPos = InStr(startPos, jsonText, ",")

    If endPos = 0 Then
        endPos = InStr(startPos, jsonText, "}")
    End If

    If endPos = 0 Then
        endPos = Len(jsonText)
    End If

    value = Trim(Mid(jsonText, startPos, endPos - startPos))
    value = Replace(value, """", "")

    ExtractJSONValue = value
End Function

'===============================================
' FUNCIÓN AUXILIAR - Codificar URL
'===============================================

Function URLEncode(textToEncode As String) As String
    '
    ' Codifica una cadena para usarla en URL
    '
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    Dim char As String

    result = ""

    For i = 1 To Len(textToEncode)
        char = Mid(textToEncode, i, 1)
        charCode = Asc(char)

        Select Case char
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", ".", "~"
                result = result & char
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Format(charCode, "00X")
        End Select
    Next i

    URLEncode = result
End Function

'===============================================
' ATAJO DE TECLADO
'===============================================

Sub Auto_Open()
    '
    ' Se ejecuta automáticamente al abrir el libro
    ' Asigna Ctrl+Shift+G a la función Geocodificar
    '
    On Error Resume Next
    Application.OnKey "+^G", "Geocodificar"
End Sub

'===============================================
' FIN DEL CÓDIGO VBA
'===============================================
