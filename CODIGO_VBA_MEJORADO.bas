'=======================================================
' GEOCODIFICADOR EXCEL - VERSIÓN MEJORADA
' Simplificado y más robusto
' Presiona Ctrl+Shift+G para ejecutar
'=======================================================

Option Explicit

Sub Geocodificar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim direccion As String
    Dim lat As String
    Dim lon As String
    Dim resultado As String
    Dim contador As Long

    On Error GoTo ErrorHandler

    ' Obtener la hoja activa
    Set ws = ThisWorkbook.Sheets("Geocodificación")

    ' Encontrar última fila con datos
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    contador = 0

    ' Procesar cada fila
    For row = 2 To lastRow
        direccion = Trim(ws.Cells(row, 1).Value)

        ' Si hay dirección, geocodificar
        If Len(direccion) > 0 Then
            ws.Cells(row, 4).Value = "⏳ Procesando..."
            ws.Cells(row, 4).Interior.Color = RGB(255, 255, 0) ' Amarillo
            DoEvents

            ' Obtener coordenadas
            If BuscarCoordenadas(direccion, lat, lon) Then
                ws.Cells(row, 2).Value = lat
                ws.Cells(row, 3).Value = lon
                ws.Cells(row, 4).Value = "✓ Encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(0, 176, 80) ' Verde
                contador = contador + 1
            Else
                ws.Cells(row, 2).Value = "N/A"
                ws.Cells(row, 3).Value = "N/A"
                ws.Cells(row, 4).Value = "✗ No encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(192, 0, 0) ' Rojo
            End If

            ' Esperar 2 segundos
            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next row

    MsgBox "✅ ¡Geocodificación completada!" & vbCrLf & _
           "Direcciones encontradas: " & contador & " de " & (lastRow - 1), _
           vbInformation, "Éxito"

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'=======================================================
' FUNCIÓN PRINCIPAL - Buscar Coordenadas
'=======================================================

Function BuscarCoordenadas(direccion As String, ByRef latitud As String, ByRef longitud As String) As Boolean
    Dim xmlhttp As Object
    Dim url As String
    Dim response As String
    Dim latPos As Long
    Dim lonPos As Long
    Dim latValue As String
    Dim lonValue As String

    On Error GoTo ErrorBuscar

    ' Crear objeto HTTP
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Construir URL
    url = "https://nominatim.openstreetmap.org/search?q=" & _
          Application.WorksheetFunction.URLEncode(direccion) & _
          "&format=json&limit=1"

    ' Hacer solicitud
    xmlhttp.Open "GET", url, False
    xmlhttp.SetRequestHeader "User-Agent", "Mozilla/5.0"
    xmlhttp.Send

    response = xmlhttp.responseText

    ' Verificar respuesta vacía
    If Len(response) = 0 Or response = "[]" Then
        BuscarCoordenadas = False
        Exit Function
    End If

    ' Extraer latitud
    latPos = InStr(response, """lat"":")
    If latPos > 0 Then
        latValue = ExtractValor(response, latPos)
        If Len(latValue) > 0 Then
            latitud = latValue
        Else
            BuscarCoordenadas = False
            Exit Function
        End If
    Else
        BuscarCoordenadas = False
        Exit Function
    End If

    ' Extraer longitud
    lonPos = InStr(response, """lon"":")
    If lonPos > 0 Then
        lonValue = ExtractValor(response, lonPos)
        If Len(lonValue) > 0 Then
            longitud = lonValue
        Else
            BuscarCoordenadas = False
            Exit Function
        End If
    Else
        BuscarCoordenadas = False
        Exit Function
    End If

    BuscarCoordenadas = True
    Exit Function

ErrorBuscar:
    BuscarCoordenadas = False
End Function

'=======================================================
' FUNCIÓN AUXILIAR - Extraer Valor JSON
'=======================================================

Function ExtractValor(jsonStr As String, startPos As Long) As String
    Dim i As Long
    Dim valor As String
    Dim charCount As Long
    Dim inQuote As Boolean

    i = startPos

    ' Buscar el primer número o comilla
    Do While i < Len(jsonStr)
        i = i + 1
        Dim char As String
        char = Mid(jsonStr, i, 1)

        If char = """" Then
            inQuote = True
            Exit Do
        ElseIf char >= "0" And char <= "9" Or char = "-" Then
            Exit Do
        End If
    Loop

    valor = ""
    charCount = 0

    ' Extraer valor
    Do While i < Len(jsonStr) And charCount < 20
        char = Mid(jsonStr, i, 1)

        If inQuote Then
            If char = """" Then Exit Do
        Else
            If char = "," Or char = "}" Or char = "]" Then Exit Do
        End If

        valor = valor & char
        i = i + 1
        charCount = charCount + 1
    Loop

    valor = Trim(Replace(valor, """", ""))
    ExtractValor = valor
End Function

'=======================================================
' ATAJO DE TECLADO
'=======================================================

Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "+^G", "Geocodificar"
End Sub
